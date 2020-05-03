#Requires -Version 5.1

Add-Type -AssemblyName System.Drawing # For Exif access
Add-Type -Path "$PSScriptRoot\VISE_MediaInfo\VISE_MediaInfo.dll"
$MediaInfo = New-Object VISE_MediaInfo.MediaInfo

Import-Module PSParallel

$ExtPhoto     = '.jpg','.jpeg','.png','.gif'
$ExtVideo     = '.3gp','.mp4','.avi','.mov','.mkv','.webp'
$Extensions   = $ExtPhoto + $ExtVideo

$NameFormats  = '(?<Year>\d{4})(?<Month>\d{2})(?<Day>\d{2})_(?<Hour>\d{2})(?<Minute>\d{2})(?<Second>\d{2})',
                '(?<Year>\d{4})-(?<Month>\d{2})-(?<Day>\d{2})_(?<Hour>\d{2})-(?<Minute>\d{2})-(?<Second>\d{2})',
                '(IMG|VID)-(?<Year>\d{4})(?<Month>\d{2})(?<Day>\d{2})-WA(?<Hour>\d{2})(?<Minute>\d{1})(?<Second>\d{1})',
                'IMG_(?<Year>\d{4})(?<Month>\d{2})(?<DayOfYear>\d{3})_(?<MinuteOfDay>\d{4})'
$NameRegex    = $NameFormats -join '|'

$NameTags     = 'NIGHT|HDR|PANO|COLLAGE|HDR-COLLAGE|EFFECTS|SCREENSHOT|EDITED|-WA\d{4}'

$DiskMutex    = New-Object System.Threading.Mutex # Don't allow parallel processing to save memory and time

#region Get dates from particular sources

function Get-VideoDate {
    param(
        [Parameter(Mandatory)]
        $FullName
    )

    try {
        $DiskMutex.WaitOne() | Out-Null
        if (!$MediaInfo.Open($FullName)) { return }
        $Raw = $MediaInfo.Get([VISE_MediaInfo.StreamKind]::General,0,'Encoded_Date',[VISE_MediaInfo.InfoKind]::Text,[VISE_MediaInfo.InfoKind]::Name)
        $MediaInfo.Close()
    }
    finally {
        $DiskMutex.ReleaseMutex() | Out-Null
    }
    
    if ($Raw) {
        [datetime]($Raw.Substring(4) + 'Z')
    }
}
function Get-ExifDate {
    param(
        [Parameter(Mandatory)]
        $FullName
    )

    $Retries       = 0
    $LastException = $false
    do {
        try {
            $DiskMutex.WaitOne() | Out-Null
            [System.Drawing.Image]$Image = [System.Drawing.Image]::FromFile($FullName)
        }
        catch [System.OutOfMemoryException] {
            $LastException = $_
            Start-Sleep -Milliseconds 500 # just wait a bit for memory situation to maybe resolve
        }
        finally {
            $DiskMutex.ReleaseMutex() | Out-Null
        }
    } while (!$Image -and $Retries -lt 3)

    if (!$Image) { throw $LastException }

    try {
        $ExifProperty = $Image.GetPropertyItem(0x9003).Value
        $DateString   = [System.Text.Encoding]::ASCII.GetString($ExifProperty)

        [DateTime]($DateString -replace '(.+?):(.+?):(.+?) ','$1.$2.$3 ')
    }
    catch {}
    finally {
        $Image.Dispose()
    }
}
function Get-MetaDate {
        param(
        [Parameter(Mandatory)]
        [string]
        $FullName
    )

    if ($File.Extension -in $ExtPhoto) {
        return Get-ExifDate $FullName
    }
    elseif ($File.Extension -in $ExtVideo) {
        return Get-VideoDate $FullName
    }
    else {
        Write-Error "Extension of $FullName is unknown"
        return
    }
}
function Get-NameDate { # add whatsapp date format - without time
    param(
        [Parameter(Mandatory)]
        $Name
    )

    if ($Name -match $NameRegex) {
        $Year  = [int]::Parse($Matches.Year)
        $Month = [int]::Parse($Matches.Month)

        if ($Matches.ContainsKey('DayOfYear'))
        {
            $Date = [datetime]::new($Year,1,1).AddDays([int]::Parse($Matches.DayOfYear) - 1).AddMinutes([int]::Parse($Matches.MinuteOfDay))
            $Day    = $Date.Day
            $Hour   = $Date.Hour
            $Minute = $Date.Minute
            $Second = 0
        }
        else
        {
            $Day    = [int]::Parse($Matches.Day)
            $Hour   = [int]::Parse($Matches.Hour)
            $Minute = [int]::Parse($Matches.Minute)
            $Second = 0
            if ($Matches.ContainsKey('Second')) { $Second = [int]::Parse($Matches.Second) }
        }

        return [datetime]::new($Year,$Month,$Day,$Hour,$Minute,$Second)
    }
}

#endregion

#region Get packaged dates

function Get-AllDates {
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]
        $FullName,

        [switch]
        $AddEarliest
    )

    process {
        $File = Get-Item $FullName

        $MetaDate = Get-MetaDate $FullName
        $NameDate = Get-NameDate $File.Name

        $Dates    = [Ordered]@{
            CreationDate     = $File.CreationTime
            ModificationDate = $File.LastWriteTime
        }
        
        if ($NameDate) { $Dates.Add('NameDate', $NameDate) }
        if ($MetaDate) { $Dates.Add('MetaDate', $MetaDate) }

        if ($AddEarliest)
        {
            $Earliest = $Dates.GetEnumerator() | Sort Value | Select -First 1
            $EarliestName = $Earliest.Name
            $EarliestDate = $Earliest.Value

            # NameDate is preferred #1
            if ($NameDate -and $EarliestName -ne 'NameDate')
            {
                if ($NameDate.Subtract($EarliestDate).TotalHours -lt 24)
                {
                    $EarliestName = 'NameDate'
                    $EarliestDate = $NameDate
                }
            }

            # MetaDate is preferred #2
            if ($MetaDate -and $EarliestName -notin 'NameDate','MetaDate')
            {
                if ($MetaDate.Subtract($EarliestDate).TotalHours -lt 24)
                {
                    $EarliestName = 'MetaDate'
                    $EarliestDate = $MetaDate
                }
            }

            $Dates = [Ordered]@{
                EarliestName = $EarliestName
                EarliestDate = $EarliestDate
            } + $Dates
        }

        $MinTimezoneStep = 30 * 60 # 15-min shift timezones are weird!
        $DiffTolerance   = 3 # seconds
        if ($NameDate -and $MetaDate)
        {
            $TimeDiff = $NameDate.Subtract($MetaDate).TotalSeconds
            $DiffByTZ = ($TimeDiff + $DiffTolerance) % $MinTimezoneStep
            if ($DiffByTZ -le $DiffTolerance*2) {
                Write-Warning "$FullName`: Probably taken in different timezone. Difference to current TZ is $([int](($TimeDiff + $DiffTolerance) / $MinTimezoneStep)) hours"
            }
        }

        [PSCustomObject] ([Ordered]@{
            FullName     = $FullName
        } + $Dates)
    }
}
function Get-BestDate {
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]
        $FullName
    )
    begin {
        function Format-Result {
            param ($Date, $DateType)

            return [PSCustomObject]@{
                FullName      = $FullName
                DirectoryName = $File.DirectoryName
                Extension     = $File.Extension
                Date          = $Date
                DateType      = $DateType
            }
        }
    }

    process {
        $File = Get-Item $FullName

        # First priority is file name as the most precise and fastest. BUT whatsapp media contains only date, so for those we also try to query Meta
        $NameDate = Get-NameDate $File.Name
        if ($NameDate -and $File.Name -notmatch '-WA\d{4}') {
            return ( Format-Result $NameDate 'NameDate' )
        }

        $MetaDate = Get-MetaDate $FullName
        if ($MetaDate) {
            return ( Format-Result $MetaDate 'MetaDate' )
        }
        # if whatsapp file contains no meta, we already have NameDate
        elseif ($File.Name -match '-WA\d{4}')
        {
            return ( Format-Result $NameDate 'NameDate' )
        }

        $Dates = @{
            CreationDate     = $File.CreationTime
            ModificationDate = $File.LastWriteTime
        }

        $Earliest = $Dates.GetEnumerator() | Sort Value | Select -First 1
        return ( Format-Result $Earliest.Value $Earliest.Name )
    }
}

#endregion

function Rename-Media {
    param(
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string]
        $FullName,

        [switch]
        $NoParallelNoProgress
    )

    process {
        $FilesToRename = Get-ChildItem $FullName -File -Recurse | ? Extension -in $Extensions | ? Name -NotMatch '^20\d{6}_\d{6}(_\w+)?\.' # (_\w)? is for _HDR, _PANO, etc

        # Collect best dates for each file
        if (!$NoParallelNoProgress) {
            $FileDates     = $FilesToRename | Invoke-Parallel -ThrottleLimit 4 { Get-BestDate $_.FullName } # We restrict parallelism with mutex. But mutex+parallel is faster than single-threaded loop.
        }
        else {
            $FileDates     = $FilesToRename | Get-BestDate
        }

        foreach ($File in $FileDates)
        {
            $Tag = ''
            if ($File.FullName -match $NameTags)
            {
                $Tag = '_' + $Matches[0].ToUpper() -replace '-WA\d{4}','WHATSAPP'
            }

            $NewName       = $File.Date.ToString('yyyyMMdd_HHmmss') + $Tag + $File.Extension
            $NameIncrement = 1
            while (Test-Path (Join-Path $File.DirectoryName $NewName)) {
                $NewName       = $File.Date.ToString('yyyyMMdd_HHmmss') + '_' + $NameIncrement + $Tag + $File.Extension
                $NameIncrement += 1
            }

            Rename-Item $File.FullName $NewName
            if ($?) { Write-Host ('{0,25} -> {1} [{2}]' -f $File.FullName, $NewName, $File.DateType ) }
        } # end foreach
    } # end process
} #end function

#Requires -Version 5.1

Add-Type -AssemblyName System.Drawing
Add-Type -Path "$PSScriptRoot\VISE_MediaInfo\VISE_MediaInfo.dll"
$MediaInfo = New-Object VISE_MediaInfo.MediaInfo

#$Global:MediaInfo = "$PSScriptRoot\MediaInfo\MediaInfo.exe"

Import-Module PSParallel

$ExtPhoto     = '.jpg','.jpeg','.png'
$ExtVideo     = '.3gp','.mp4','.avi','.mov','.mkv'
$Extensions   = $ExtPhoto + $ExtVideo

$NameFormats  = '(?<Year>\d{4})(?<Month>\d{2})(?<Day>\d{2})_(?<Hour>\d{2})(?<Minute>\d{2})(?<Second>\d{2})',
                'IMG_(?<Year>\d{4})(?<Month>\d{2})(?<DayOfYear>\d{3})_(?<MinuteOfDay>\d{4})'
$NameRegex    = $NameFormats -join '|'

$VideoMutex   = New-Object System.Threading.Mutex # Don't allow video parallel processing to save memory and time

function Get-VideoDate {
    param(
        [Parameter(Mandatory)]
        $Path
    )

    try {
        $VideoMutex.WaitOne() | Out-Null
        if ($MediaInfo.Open($Path)) {
            $Raw = $MediaInfo.Get([VISE_MediaInfo.StreamKind]::General,0,'Encoded_Date',[VISE_MediaInfo.InfoKind]::Text,[VISE_MediaInfo.InfoKind]::Name)
            $MediaInfo.Close()

            return [datetime]($Raw.Substring(4) + 'Z')
        }
    }
    finally {
        $VideoMutex.ReleaseMutex() | Out-Null
    }
}

<#function Get-VideoDate {
    param(
        [Parameter(Mandatory)]
        $Path
    )
    try {
        $VideoMutex.WaitOne()
        (& $Global:MediaInfo $Path | sls 'encoded.+?UTC (.+)').Matches[0].Groups[1].Value.ToDateTime($null)
    }
    finally {
        $VideoMutex.ReleaseMutex()
    }
}#>

function Get-ExifDate {
    param(
        [Parameter(Mandatory)]
        $Path
    )

    $Retries       = 0
    $LastException = $false
    do {
        try {
            [System.Drawing.Image]$Image = [System.Drawing.Image]::FromFile($Path)
        }
        catch [System.OutOfMemoryException] {
            $LastException = $_
            Start-Sleep -Milliseconds 500 # just wait a bit for memory situation to maybe resolve
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

function Get-NameDate {
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

# TODO: Read Exif geotag for timezone-aware comparison of Exif <> Name dates
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

function Rename-Media {
    param(
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string]
        $FullName
    )

    process {
        $FilesToRename = Get-ChildItem $FullName -File -Recurse | ? Extension -in $Extensions | ? Name -NotMatch '^20\d{6}_\d{6}(_\w+)?\.' # (_\w)? is for _HDR, _PANO, etc
        $FilesToRename = $FilesToRename | Get-Random -Count $FilesToRename.Count # Randomization is essential for videos to be evenly mixed with photos (video processing is single-threaded)
        $FileDates     = $FilesToRename | Invoke-Parallel -ThrottleLimit 64 { Get-AllDates $_.FullName }
        foreach ($File in $FileDates)
        {
            
        } # end foreach
    } # end process
} #end function

function Rename-Media_Old {
    param(
        [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string]
        $FullName
    )

    process {
        $FilesToRename = Get-ChildItem $FullName -File -Recurse | ? Extension -in $Extensions | ? Name -NotMatch '^20\d{6}_\d{6}(_\w+)?\.'
        foreach ($File in $FilesToRename)
        {
            if ($File.Name -match $NameRegex)
            {
                $Year  = [int]::Parse($Matches.Year)
                $Month = [int]::Parse($Matches.Month)

                if ($Matches.ContainsKey('DayOfYear'))
                {
                    $Date = [datetime]::new($Year,1,1).AddDays([int]::Parse($Matches.DayOfYear) - 1).AddMinutes([int]::Parse($Matches.MinuteOfDay))
                    if ($Date.Month -ne $Month)
                    {
                        Write-Warning "Date not consistent for $($File.Name)"
                        $CancelRename = $true
                    }
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

                if (!$CancelRename)
                {
                    $NewName       = ('{0:0000}{1:00}{2:00}_{3:00}{4:00}{5:00}' -f $Year,$Month,$Day,$Hour,$Minute,$Second) + $File.Extension
                    $NameIncrement = 1
                    while (Test-Path (Join-Path $File.DirectoryName $NewName)) {
                        $NewName       = ('{0:0000}{1:00}{2:00}_{3:00}{4:00}{5:00}_{6}' -f $Year,$Month,$Day,$Hour,$Minute,$Second,$NameIncrement) + $File.Extension
                        $NameIncrement += 1
                    }

                    Rename-Item $File.FullName $NewName
                    if ($?) { Write-Host ('{0,25} -> {1}' -f $File.Name, $NewName ) }
                }
            }
            $CancelRename = $false
        } # end foreach
    } # end process
} # end function
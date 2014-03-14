[CmdletBinding()]
Param (
    [ValidateScript({Test-Path $_})]
    [string] $FitFile,
    [string] $HrvLog = $null,
    [switch] $ConvertPosition)
    
Set-StrictMode -Version 2.0

$dllPath = Join-Path ([IO.Path]::GetDirectoryName($PSCommandPath)) "Fit.dll"
$null = [Reflection.Assembly]::LoadFrom($dllPath)

Function ConvertGarminProductName($id)
{
    $fields = [Dynastream.Fit.GarminProduct].GetFields() | Select -Property Name
    foreach ($f in $fields) {
        $name = $f.Name
        Invoke-Expression "`$value = [Dynastream.Fit.GarminProduct]::$name"
        if ($value -eq $id) {
            return $name
        }
    }
    
    return "UnknownProduct"
}

Function ConvertManufacturerName($id)
{
    $fields = [Dynastream.Fit.Manufacturer].GetFields() | Select -Property Name
    foreach ($f in $fields) {
        $name = $f.Name
        Invoke-Expression "`$value = [Dynastream.Fit.Manufacturer]::$name"
        if ($value -eq $id) {
            return $name
        }
    }
    
    return "UnknownManufacturer"
}

$global:g_activity = New-Object PSObject
$global:g_sessions = @()
$global:g_laps = @()
$global:g_records = @()
$global:g_recoveryHr = 0
# coefficient to convert GPS location from semicircle unit to degree.
$global:g_semicircle2degree = 180.0 / [Math]::Pow(2, 31)

# List of known messages that we recognize and unknown ones that we explicitly do not.
$global:g_knownMesgs = @{
    0 = "FileId";
    1 = "Capabilities";
    18 = "Session";
    19 = "Lap";
    20 = "Record";
    21 = "Event";
    22 = "unknown";
    23 = "DeviceInfo";
    29 = "unknown";
    34 = "Activity";
    35 = "Software";
    37 = "FileCapabilities";
    38 = "MesgCapabilities";
    39 = "FieldCapabilities";
    49 = "FileCreator";
    78 = "Hrv";
    79 = "unknown";
    104 = "unknown";
}

# Action for processing unrecognized messages
$onMesg = {
    $mesg = $EventArgs.mesg
    $num = [int] $mesg.Num
    
    if ($global:g_knownMesgs.ContainsKey($num)) {
        return
    }
    
    Write-Warning "OnMesg $num $($mesg.Name)"
    $fieldCount = $mesg.GetNumFields()

    # Prints all the fields at the Verbose mode.
    for ($i = 0; $i -lt $fieldCount; $i++) {
        $field = $mesg.fields[$i]
        $fieldNum = $field.Num
        $fieldName = $field.GetName()
        $valueCount = $field.GetNumValues()

        Write-Verbose "Field $fieldName $valueCount"

        for ($j = 0; $j -lt $valueCount; $j++) {
            Write-Verbose "  => $($field.GetValue($j)) $fieldNum RawValue=$($field.GetRawValue($j))"
        }
    }
}

$decoder = New-Object Dynastream.Fit.Decode
$global:mesgBroadcaster = New-Object Dynastream.Fit.MesgBroadcaster

$decoder.add_MesgEvent({
    Param ($Sender, $EventArgs)

    $global:mesgBroadcaster.OnMesg($Sender, $EventArgs)
})

$decoder.add_MesgDefinitionEvent({
    Param ($Sender, $EventArgs)

    $def = $EventArgs.mesgDef
    Write-Verbose "MesgDefinition: $($def.LocalMesgNum) $($def.GlobalMesgNum) NumFields = $($def.NumFields) MesgSize = $($def.GetMesgSize())"
})

$mesgBroadcaster.add_MesgEvent($onMesg)

$mesgBroadcaster.add_ActivityMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    Add-Member -InputObject $global:g_activity NoteProperty Type $mesg.GetType()
    Add-Member -InputObject $global:g_activity NoteProperty Event $mesg.GetEvent()
    Add-Member -InputObject $global:g_activity NoteProperty EventType $mesg.GetEventType()
    Add-Member -InputObject $global:g_activity NoteProperty Timestamp $mesg.GetTimestamp().GetDateTime()
    Add-Member -InputObject $global:g_activity NoteProperty TotalTimerTime $mesg.GetTotalTimerTime()
    $numSessions = $mesg.GetNumSessions()
    Write-Host "Activity: Sessions = $numSessions"
})

$mesgBroadcaster.add_LapMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $fieldCount = $mesg.GetNumFields()
    
    $lap = New-Object PSObject

    for ($i = 0; $i -lt $fieldCount; $i++) {
        $field = $mesg.fields[$i]
        $fieldName = $field.GetName()
        $valueCount = $field.GetNumValues()

        for ($j = 0; $j -lt $valueCount; $j++) {
            $value = $field.GetValue($j)

            if ($fieldName -eq "StartTime" -or $fieldName -eq "Timestamp") {
                $value = $mesg.TimestampToDateTime($value).GetDateTime()
            }
            elseif ($ConvertPosition -and $fieldName.Contains("PositionL")) {
                $value *= $global:g_semicircle2degree
            }

            if ($fieldName -ne "unknown") {
                Add-Member -InputObject $lap NoteProperty $fieldName $value
            }
        }
    }
    
    Add-Member -InputObject $lap NoteProperty Records $global:g_records
    $global:g_records = @()
    
    $global:g_laps += $lap
})

$mesgBroadcaster.add_RecordMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $fieldCount = $mesg.GetNumFields()
    
    $record = New-Object PSObject

    for ($i = 0; $i -lt $fieldCount; $i++) {
        $field = $mesg.fields[$i]
        $fieldName = $field.GetName()
        $valueCount = $mesg.fields[$i].GetNumValues()

        for ($j = 0; $j -lt $valueCount; $j++) {
            $value = $field.GetValue($j)

            if ($fieldName -eq "StartTime" -or $fieldName -eq "Timestamp") {
                $value = $mesg.TimestampToDateTime($value).GetDateTime()
            }
            elseif ($ConvertPosition -and $fieldName.StartsWith("PositionL")) {
                $value *= $global:g_semicircle2degree
            }
            
            Add-Member -InputObject $record NoteProperty $fieldName $value
        }
    }
    $global:g_records += $record
    
    $hr = $mesg.GetHeartRate()
    if ($hr -ne $null -and (-not [String]::IsNullOrEmpty($HrvLog))) {
        "{0}, {1}, " -f $record.Timestamp, $hr | Out-File -Append -FilePath $HrvLog -Encoding ASCII
    }
})

$mesgBroadcaster.add_FileIdMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $fileType = $mesg.GetType()
    $manufacturer = ConvertManufacturerName $mesg.GetManufacturer()
    $product = ConvertGarminProductName $mesg.GetProduct()
    $serialNumber = $mesg.GetSerialNumber()
    $number = $mesg.GetNumber()
    $timeCreated = $mesg.GetTimeCreated().GetDateTime()
    
    Write-Host "File type = $fileType, $manufacturer $product, Serial Number $serialNumber $number.  Created $timeCreated"
})

$mesgBroadcaster.add_UserProfileMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $friendlyName = [System.Text.Encoding]::UTF8.GetString($mesg.GetFriendlyName())
    $gender = $mesg.GetGender().ToString()
    $age = $mesg.GetAge()
    $weight = $mesg.GetWeight()
    $mesg | Format-List | Out-String | Write-Host
    
    Write-Host "UserProfile: Name $friendlyName $gender Age = $age Weight = $weight"
})

$mesgBroadcaster.add_EventMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $event = $mesg.GetEvent()
    $eventType = $mesg.GetEventType()
    $eventGroup = $mesg.GetEventGroup()
    $data = $mesg.GetData()
    $time = $mesg.GetTimestamp().GetDateTime()
    Write-Host "$time : $eventType $event event group = $eventGroup data = $data"

    if ($eventType -eq "Marker" -and $event -eq "RecoveryHr") {
        $global:g_recoveryHr = $data
    }
})

$mesgBroadcaster.add_FileCreatorMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $softwareVer = $mesg.GetSoftwareVersion()
    $hardwareVer = $mesg.GetHardwareVersion()
    Write-Host "Software version $softwareVer Hardware version $hardwareVer"
})

$mesgBroadcaster.add_CapabilitiesMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $workoutsSupported = $mesg.GetWorkoutsSupported()
    Write-Host "Capabilities: $workoutsSupported"
})

$mesgBroadcaster.add_FileCapabilitiesMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $type = $mesg.GetType()
    $dir = [System.Text.Encoding]::UTF8.GetString($mesg.GetDirectory())
    Write-Host "FileCapabilities: $type $dir"
})

$mesgBroadcaster.add_FieldCapabilitiesMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $file = $mesg.GetFile()
    Write-Host "FieldCapabilities: $file"
})

$mesgBroadcaster.add_MesgCapabilitiesMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $file = $mesg.GetFile()
    Write-Host "MesgCapabilities: $file"
})

$mesgBroadcaster.add_SoftwareMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $ver = $mesg.GetVersion()
    $partNumber = [System.Text.Encoding]::UTF8.GetString($mesg.GetPartNumber())
    Write-Host "Software: $ver $partNumber"
})

$mesgBroadcaster.add_HrvMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $hrv = 0..($mesg.GetNumTime() - 1) | % {
        $value = $mesg.GetTime($_)
        if ($value -lt 65.534) {
            if (-not [String]::IsNullOrEmpty($HrvLog)) {
                ", , {0}" -f $value | Out-File -Append -FilePath $HrvLog -Encoding ASCII
            }
        }
    }
})

$mesgBroadcaster.add_SessionMesgEvent({
    Param ($Sender, $EventArgs)

    $mesg = $EventArgs.mesg
    $fieldCount = $mesg.GetNumFields()
    $session = New-Object PSObject

    for ($i = 0; $i -lt $fieldCount; $i++) {
        $field = $mesg.fields[$i]
        $fieldName = $field.GetName()
        $valueCount = $field.GetNumValues()

        for ($j = 0; $j -lt $valueCount; $j++) {
            $value = $field.GetValue($j)

            if ($fieldName -eq "StartTime" -or $fieldName -eq "Timestamp") {
                $value = $mesg.TimestampToDateTime($value).GetDateTime()
            }
            elseif ($ConvertPosition -and ($fieldName.Contains("PositionL") -or $fieldName.EndsWith("cLat") -or $fieldName.EndsWith("cLong"))) {
                $value *= $global:g_semicircle2degree
            }

            if ($fieldName -ne "unknown") {
                Add-Member -InputObject $session NoteProperty $fieldName $value
            }
        }
    }
    
    $session.Event = $mesg.GetEvent()
    $session.EventType = $mesg.GetEventType()
    $session.Sport = $mesg.GetSport()
    $session.SubSport = $mesg.GetSubSport()
    $session.Trigger = $mesg.GetTrigger()

    Add-Member -InputObject $session NoteProperty RecoveryHeartRate $global:g_recoveryHr

    Add-Member -InputObject $session NoteProperty Laps $global:g_laps
    $global:g_laps = @()
    
    $global:g_sessions += $session
})

# Reads and processes the file

$fitSource = New-Object System.IO.FileStream($FitFile, [IO.FileMode]::Open, [IO.FileAccess]::Read)
if ($decoder.IsFit($fitSource)) {
    if (-Not $decoder.CheckIntegrity($fitSource)) {
        Write-Warning "FIT file integrity check failed"
    }
    
    try {
        if (-not [String]::IsNullOrEmpty($HrvLog)) {
            "Time, HR, R-R" | Out-File -FilePath $HrvLog -Encoding ASCII
        }
        
        if (-Not $decoder.Read($fitSource)) {
            Write-Warning "Failed to read the FIT file"
        }
    }
    finally {
        $fitSource.Close()
    }
}

# Shows the results

$global:g_sessions.Laps | `
    Select -Property Timestamp, StartTime, TotalElapsedTime, TotalDistance, AvgHeartRate, MaxHeartRate, AvgCadence, MaxCadence | `
    Sort -Property Timestamp | Format-Table | Out-String | Write-Host

# Generates the activity object
Add-Member -InputObject $global:g_activity NoteProperty Sessions $global:g_sessions
return $global:g_activity

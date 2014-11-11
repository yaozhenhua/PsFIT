Function GetFitTimestamp($fileName)
{
    $dllPath = Join-Path ([IO.Path]::GetDirectoryName($PSCommandPath)) "Fit.dll"
    $null = [Reflection.Assembly]::LoadFrom($dllPath)

    $decoder = New-Object Dynastream.Fit.Decode
    $decoder.add_MesgEvent({
        Param ($Sender, $EventArgs)

        $global:mesgBroadcaster.OnMesg($Sender, $EventArgs)
    })

    $global:mesgBroadcaster = New-Object Dynastream.Fit.MesgBroadcaster

    $global:fitLapTimestamp = @()
    $mesgBroadcaster.add_LapMesgEvent({
            param ($Sender, $EventArgs)

            $mesg = $EventArgs.mesg
            $global:fitLapTimestamp += $mesg.GetStartTime().GetDateTime()
        })

    $fitSource = New-Object System.IO.FileStream($fileName, [IO.FileMode]::Open, [IO.FileAccess]::Read)
    if ($decoder.IsFit($fitSource)) {
        try {
            if ($decoder.CheckIntegrity($fitSource) -and $decoder.Read($fitSource)) {
                $startTime = $global:fitLapTimestamp[0]
                # Convert to local time
                return [TimeZoneInfo]::ConvertTimeFromUtc($startTime, [TimeZoneInfo]::Local)
            }
        }
        finally {
            $fitSource.Close()
        }
    }

    return $null
}

GetFitTimestamp c:\users\zhyao\AppData\Roaming\Garmin\Devices\3855590556\Activities\20140708-065603-1-1328-ANTFS-4-0.FIT

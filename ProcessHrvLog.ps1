Param ($HrvLog = "hrv.log")

$time = $null
$hr = $null

$content = Get-Content $HrvLog
foreach ($line in $content) {
    if ($line -notmatch "Time, HR, R-R") {
        $elements = $line -split ",\s*"
        if ([string]::IsNullOrEmpty($elements[0])) {
            $hrv = $elements[2]
            $dev = [float] $hrv / (60.0 / $hr) - 1.0
            if ($dev -le 0.5) {
                New-Object PSObject -Property @{
                    Timestamp = $time;
                    HR = $hr;
                    HRV = $hrv;
                    Deviation = $dev;
                }
            }
        }
        else {
            $time = [DateTime] $elements[0]
            $hr = [int] $elements[1]
        }
    }
}

$HrvFile = "C:\Users\zhyao\Desktop\hrv.csv"

# Max RR interval should be the avg RR multiplies the following factor.  It must be greater than 1.
$m_rrErrorRange = 2.0
# How long the moving average window is.  It should be in the range of (1, 5).
$m_windowInMinutes = 1

Set-StrictMode -Version latest

Function CalculateRMSSD([Collections.ArrayList]$movingWindow)
{
    if ($movingWindow.Count -eq 0) {
        return
    }

    $startTime = $movingWindow[-1].Time - [TimeSpan]::FromMinutes($m_windowInMinutes)
    while ($movingWindow[0].Time -lt $startTime) {
        $movingWindow.RemoveAt(0)
    }

    $n = 0
    $sum = 0.0
    $hr = 0
    $lastRR = 0
    foreach ($point in $movingWindow) {
        # Base line R-R interval for this heart rate, unit in second.
        $baselineRR = 60.0 / $point.HR

        $point.RR | % {
            # Filter out the bad data (missing signal causes R-R interval too long)
            if ($_ -lt $baselineRR * $m_rrErrorRange) {
                if ($lastRR -gt 0) {
                    $n++
                    $d = ($_ - $lastRR) # / $baselineRR
                    $sum += $d * $d
                }

                $lastRR = $_
            }
        }

        $hr += $point.HR
    }

    # Average heart rate
    $hr = $hr / $movingWindow.Count
    # Change unit from second to millisecond
    $rMSSD = [Math]::Sqrt($sum / $n) * 1000.0

    return (New-Object PSObject -Property @{
        Time = $movingWindow[-1].Time
        HR_avg = $hr
        # Firstly change the RMS unit from second to millisecond, then take 20*Ln(x) to normalize the data.
        # Reference:
        #   http://hrvtraining.com/2013/07/04/rmssd-the-hrv-value-provided-by-ithlete-and-bioforce/
        #   http://hrvtraining.com/2013/11/22/hrv-in-a-bit-more-detail-part-2/
        HRV = 20.0 * [Math]::Log($rMSSD)
        rMSSD = $rMSSD
    })
}

# Import the HRV data, three columns are Time, HR, and R-R (in second)
$hrvData = Import-Csv $HrvFile

# Data structure:
#   ( (Time, HR, (R-R, R-R, ...), (Time, HR, (R-R, R-R, ...), ...)
$movingWindow = New-Object Collections.ArrayList

$lastCalculateTime = [DateTime]::MinValue
$lastTime = $null
$lastHr = $null
$rr = @()
foreach ($item in $hrvData) {
    if ([string]::IsNullOrEmpty($item.Time)) {
        $rr += [double] $item.'R-R'
    }
    else {
        if ($rr.Count -gt 0) {
            $point = New-Object PSObject -Property @{
                Time = $lastTime
                HR = $lastHr
                RR = $rr
            }
            [void] $movingWindow.Add($point)

            if (($movingWindow[-1].Time - $movingWindow[0].Time).TotalMinutes -ge $m_windowInMinutes) {
                if (($movingWindow[-1].Time - $lastCalculateTime).TotalMinutes -ge 0.5 * $m_windowInMinutes) {
                    CalculateRMSSD $movingWindow
                    $lastCalculateTime = $movingWindow[-1].Time
                }
            }
        }
        
        $lastTime = [DateTime]$item.Time
        $lastHr = [int] $item.HR
        $rr = @()
    }
}

[CmdletBinding(SupportsShouldProcess=$false)]
Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Activity)

Set-StrictMode -Version 2.0

[double] $mile = 1609.344

$lastTime = $null
$lastSpeed = $null
$totalDistance = 0.0
$totalDistanceFP = 0.0

foreach ($lap in $Activity.Sessions[0].Laps) {
    [double] $totalMeters = $lap.TotalDistance
    [double] $totalSeconds = $lap.TotalTimerTime
    $avgSpeed = $totalMeters / $totalSeconds
            
    $pace = [TimeSpan]::FromSeconds($mile / $avgSpeed).ToString("mm\:ss\.fff")
            
    $totalDistance += $totalMeters

    $lapDistanceFP = 0.0
    foreach ($record in $lap.Records) {
        if ($record.Speed -eq $null) {
            continue
        }

        $time = $record.Timestamp
        $speed = [double] $record.Speed
        # Distance in the record is ignored here.

        if ($lastTime -ne $null) {
            $delta = ($time - $lastTime).TotalSeconds * ($lastSpeed + $speed) * 0.5
            $lapDistanceFP += $delta
            $totalDistanceFP += $delta
        }
        $lastTime = $time
        $lastSpeed = $speed
    }

    $scale = $totalDistance / $totalDistanceFP
    $diff = ($scale - 1.0).ToString("P1")

    $avgCadence = $lap.AvgCadence
    $avgSpeed = $lap.AvgSpeed
    $avgHR = $lap.AvgHeartRate

    "  LAP distance {0,6} Foot pod distance {1,6} Scale {2:0.000} {3,6} Cadence {4:###} HR {5:###} Pace {6} Eff {7:0.00}" `
        -f $totalMeters.ToString("F1"), $lapDistanceFP.ToString("F1"), `
        $scale, $diff, $avgCadence, $avgHR, $pace, ($avgSpeed * 100.0 / $avgHR)
}

$scale = $totalDistance / $totalDistanceFP
$diff = ($scale - 1.0).ToString("P1")
""
"  SUM distance {0,6} Foot pod distance {1,6} Scale {2:0.000} {3,6}" `
-f $totalDistance.ToString("F1"), $totalDistanceFP.ToString("F1"), $scale, $diff
""

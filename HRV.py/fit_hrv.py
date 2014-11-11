#!/usr/bin/env python

import json
import math
import sys

# fitparse: https://github.com/dtcooper/python-fitparse/tree/ng
from fitparse import FitFile

class _HrvPoint:
    def __init__(self, _timestamp, _hr):
        self.timestamp = _timestamp
        self.hr = _hr
        self.rr = []

    def add_hrv(self, data):
        self.rr.extend(filter(lambda x: x != None, data))

    def has_rr(self):
        return len(self.rr) > 0

    def __str__(self):
        return str(self.timestamp) + str(self.hr) + str(self.rr)

def datapoints_duration_in_minutes(datapoints):
    if len(datapoints) < 2:
        return 0

    timediff = datapoints[-1].timestamp - datapoints[0].timestamp
    return timediff.total_seconds() / 60.0

def relative_error(x_0, x):
    return abs(x_0 - x) / x_0

def rmssd(datapoints, duration_in_minutes):
    # If an R-R is out of the average interval by 50%, it is a measurement error, i.e. 2 or more adjacent interval
    # merged into one by mistake. In this case we should ignore the error.
    max_relative_error = 0.50

    while True:
        if datapoints_duration_in_minutes(datapoints) > duration_in_minutes:
            del datapoints[0]
        else:
            break

    if len(datapoints) < 2:
        return None

    avg_hr = 0.0
    # Calculation of unbiased sample standard deviation.  Reference: http://en.wikipedia.org/wiki/Standard_deviation
    sum_err_2 = 0.0
    rrs = []
    for datapoint in datapoints:
        avg_hr += datapoint.hr
        baseline_rr = 60.0 / datapoint.hr * 1000.0

        # Calulate the stddev first
        for rr in datapoint.rr:
            if relative_error(baseline_rr, rr) > max_relative_error:
                continue
            rrs.append((rr, baseline_rr))
            sum_err_2 += (rr - baseline_rr) ** 2

    sigma3 = math.sqrt(sum_err_2 / (len(rrs) - 1.5)) * 3.0

    sum_rr_squre = 0.0
    num_rr = 0
    last_rr = 0
    delta = 0
    for (rr, baseline_rr) in rrs:
        if abs(rr - baseline_rr) > sigma3:
            continue

        if last_rr > 0:
            sum_rr_squre += (rr - last_rr) * (rr - last_rr)
            num_rr += 1

        last_rr = rr

        if abs(rr - baseline_rr) > delta:
            delta = abs(rr - baseline_rr)

    result = math.sqrt(sum_rr_squre / num_rr)
    hrv = math.log(result) * 20
    avg_hr /= len(datapoints)
    print "{0} HR: {1:.1f} rMSSD: {2:.3f} HRV: {3:.3f} |d_RR|_max: {4}".format(
            datapoints[-1].timestamp, avg_hr, result, hrv, delta)
    return (datapoints[-1].timestamp, avg_hr, result, hrv)

def main():
    if len(sys.argv) == 3:
        filename = sys.argv[1]
        output_filename = sys.argv[2]
    else:
        print 'Usage: {0} [FIT input file] [HTML output file]'.format(sys.argv[0])
        return

    moving_window = 1
    fit = FitFile(filename)
    fit.parse()

    # Gets the start timestamp
    start_time = None
    for message in fit.get_messages(name = 'record'):
        start_time = message.get_value('timestamp')
        break

    last_rmssd_time = None
    hrv_points = []
    datapoint = None
    hrv_results = [["Duration", "Avg HR", "HRV"]]
    for message in fit.messages:
        if message.mesg_num == 20:
            if datapoint != None and datapoint.has_rr():
                hrv_points.append(datapoint)

                if datapoints_duration_in_minutes(hrv_points) > moving_window:
                    if last_rmssd_time == None or (hrv_points[-1].timestamp - last_rmssd_time).total_seconds() / 60 > 0.5 * moving_window:
                        last_rmssd_time = hrv_points[-1].timestamp
                        result = rmssd(hrv_points, moving_window)
                        if result != None:
                            hrv_results.append([(result[0] - start_time).total_seconds(), result[1], result[2]])
    
            datapoint = _HrvPoint(message.get_value('timestamp'), message.get_value('heart_rate'))
        elif message.mesg_num == 78:
            if datapoint != None:
                datapoint.add_hrv(message.get_value('time'))
        elif message.name == 'event':
            print '{0} Event: {1} {2} {3}'.format(
                    message.get_value('timestamp'), message.get_value('event'), message.get_value('event_type'),
                    message.get_value('data'))
        elif message.name in ['session', 'lap']:
            print '{0} {1}'.format(message.get_value('timestamp'), message.name)
            for f in message.fields:
                if f.value is None or f.name in ['timestamp']:
                    continue
                print '        {0} : {1}'.format(f.name, f.value)
    
    with open('fit_hrv_template.html', 'r') as template:
        with open(output_filename, 'w') as output:
            output.write(template.read().replace("%HRVDATA%", json.dumps(hrv_results)))

if __name__ == '__main__':
    main()

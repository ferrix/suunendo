import os
import re
import sys
from glob import glob

import pytz
from lxml import etree
from dateutil import tz
from datetime import datetime
from xlrd import open_workbook

timezone = pytz.timezone('EET')

def find_summary_index(sheet):
    for i in range(30):
        if sheet.cell(0, i).value == 'Summary Fields':
            return i
    raise KeyError('Summary Fields not found')

def find_move_samples_index(sheet, summary_field_index=0):
    for i in range(summary_field_index, min(sheet.ncols, 50)):
        if sheet.cell(0, i).value == 'Move samples':
            return i
    return False

def get_basic_dict(sheet, stop_index):
    basic = {}
    for i in range(stop_index):
        key = str(sheet.cell(1, i).value)
        key = re.sub(' \[.*', '', key)
        if key == 'Tags':
            basic[key] = []
            for j in range(2, min(sheet.ncols, 30)):
                try:
                    value = sheet.cell(j, i).value
                    if value:
                        basic[key].append(sheet.cell(j, i).value)
                    else:
                        break
                except:
                    break
        else:
            basic[key] = sheet.cell(2, i).value

    return basic

def convert_xlsx_to_tcx(filename, activities):
    activity = etree.SubElement(activities, ns+'Activity')
    lap = etree.SubElement(activity, ns+'Lap')
    intensity = etree.SubElement(lap, ns+'Intensity')
    intensity.text = 'Active'

    o = open_workbook(filename)
    dest = os.path.splitext(os.path.basename(filename))[0] + '.tcx'
    sheet = o.sheets()[0]
    summary_index = find_summary_index(sheet)
    sample_index = find_move_samples_index(sheet, summary_index)
    basic = get_basic_dict(sheet, summary_index)

    if 'StartTime' in basic:
        start = datetime.strptime(basic['StartTime'], '%Y-%m-%d %H:%M:%S')
        start = timezone.localize(start)
        act_id = etree.SubElement(activity, ns+'Id')
        act_id.text = start.astimezone(tz.gettz('UTC')).strftime("%Y-%m-%dT%H:%M:%SZ")
        lap.set('StartTime', act_id.text)
    if 'Activity' in basic:
        sport = basic['Activity']
        activity.set('Sport', str(basic['Activity']))
    if 'Device' in basic:
        creator = etree.SubElement(activity, ns+'Creator')
        creator.text = str(basic['Device'])
    notes_text = []
    if 'Notes' in basic:
        notes_text.append(unicode(basic['Notes']))
    if 'Tags' in basic:
        notes_text.append("Tags:\n")
        notes_text.append(unicode("\n".join(basic['Tags'])))
    if notes_text:
        notes = etree.SubElement(activity, ns+'Notes')
        notes.text = "\n".join(notes_text)
    if 'Duration' in basic:
        duration = etree.SubElement(lap, ns+'TotalTimeSeconds')
        duration.text = str(float(basic['Duration']))
    if 'Calories' in basic:
        calories = etree.SubElement(lap, ns+'Calories')
        calories.text = str(int(basic['Calories']))
    if 'Distance' in basic:
        distance = etree.SubElement(lap, ns+'DistanceMeters')
        distance.text = str(float(basic['Distance']))
    if 'HrAvg' in basic:
        avghr = etree.SubElement(lap, ns+'AverageHeartRateBpm')
        avghr.text = str(int(basic['HrAvg']))
    if 'HrPeak' in basic:
        maxhr = etree.SubElement(lap, ns+'MaximumHeartRateBpm')
        maxhr.text = str(int(basic['HrPeak']))
    if 'SpeedMax' in basic:
        maxspeed = etree.SubElement(lap, ns+'MaximumSpeed')
        maxspeed.text = str(float(basic['SpeedMax']))

    track = etree.SubElement(lap, ns+'Track')

    if sample_index:
        for row in range(2,sheet.nrows):
            timestamp = sheet.cell(row, sample_index).value
            if timestamp:
                trackpoint = etree.SubElement(track, 'Trackpoint')
                hrmdate = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S")
                hrmdate = timezone.localize(hrmdate)
                endotime = etree.SubElement(trackpoint, 'Time')
                endotime.text = hrmdate.astimezone(tz.gettz('UTC')).strftime("%Y-%m-%dT%H:%M:%SZ")
                endohrm = etree.SubElement(trackpoint, 'HeartRateBpm')
                endohrm.text = str(int(sheet.cell(row, sample_index+2).value))

if __name__ == '__main__':
    if 2 > len(sys.argv):
        print "Filename(s) required"
        sys.exit(1)

    root_name = 'TrainingCenterDatabase'
    ns = 'http://www.garmin.com/xmlschemas/TrainingCenterDatabase/v2'
    nsmap = {None: ns}
    ns = '{' + ns + '}'
    root = etree.Element(ns+root_name, nsmap=nsmap)
    activities = etree.SubElement(root, ns+'Activities')

    for file in sys.argv[1:]:
        for name in glob(file):
            try:
                convert_xlsx_to_tcx(name, activities)
            except KeyError:
                print name, 'not converted'

    output = open('conversion.tcx',"wb")
    output.write(etree.tostring(root, pretty_print=True))
    output.close()


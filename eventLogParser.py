import win32evtlog
from datetime import datetime
from collections import deque
from time import mktime
import csv
import os


def dayMonthYearSum(date_object):
    return date_object.day + date_object.month + date_object.year

log_file_path = r"C:\Users\Sebastian Kieritz\Documents\WorkLog"
log_file_name = r"EventList-{}-{:0>2}.csv"
# default date for last run
last_run = datetime(2017, 4, 1, 0, 0, 0)

# get the date and time of the last run
for file_name in sorted(os.listdir(log_file_path), reverse=True):
    if file_name.startswith("EventList-") and file_name.endswith(".csv"):
        with open(os.path.join(log_file_path, file_name)) as csv_file:
            csv_deque = deque(csv.DictReader(csv_file), 1)
            if csv_deque:
                last_run = datetime.strptime(csv_deque[0]["TimeGenerated"], "%Y-%m-%d %H:%M:%S")
                break

print('last run:', last_run)

# iterate through the events, beginning with the most recent,
# and save the events of interest (lock- and unlock-events, as
# well as the first and last event of each day).
event_logs_of_interest = ['Application', 'System', 'Security']
all_events = []

for event_log_name in event_logs_of_interest:
    print('Processing {} event log'.format(event_log_name))
    last_event = None
    continue_while = True
    
    event_log = win32evtlog.OpenEventLog(None, event_log_name)
    while continue_while:
        events = win32evtlog.ReadEventLog(event_log, win32evtlog.EVENTLOG_BACKWARDS_READ|win32evtlog.EVENTLOG_SEQUENTIAL_READ, 0)
        if not events:
            break

        for event in events:
            event_timetuple = event.TimeGenerated.timetuple()
            event_timestamp = mktime(event_timetuple)
            event_datetime = datetime.fromtimestamp(event_timestamp)

            if event_datetime <= last_run:
                continue_while = False
                break
            else:
                all_events.append(event)

    win32evtlog.CloseEventLog(event_log)
    event_log = None

all_events = sorted(all_events, key=lambda e: e.TimeGenerated.timetuple(), reverse=True)
events_of_interest = []
last_event = None
for event in all_events:
    event_timetuple = event.TimeGenerated.timetuple()
    event_timestamp = mktime(event_timetuple)
    event_datetime = datetime.fromtimestamp(event_timestamp)

    if event_datetime <= last_run:
        continue_while = False
        break

    if last_event and dayMonthYearSum(event.TimeGenerated) != dayMonthYearSum(last_event.TimeGenerated):
        # the day changed, append the current and perhaps the old event
        if not last_event_was_appended:
            events_of_interest.append(last_event)
        events_of_interest.append(event)
        last_event_was_appended = True
    elif event.EventID in (4800, 4801): # append events of interest: 4800 = lock, 4801 = unlock
        events_of_interest.append(event)
        last_event_was_appended = True
    else:
        last_event_was_appended = False
    last_event = event

# save last event if it was not yet appended and the day differs from
# the day of the last run (hence the last run must have been the last event of that day)
if last_event and not last_event_was_appended and dayMonthYearSum(last_event.TimeGenerated) != dayMonthYearSum(last_run):
    events_of_interest.append(last_event)

print()
# write all events of interest to the corresponding csv file
last_written = last_run
event_index = 0
events_of_interest.reverse()
while event_index < len(events_of_interest):
    current_log_file_name = log_file_name.format(last_written.year, last_written.month)
    complete_path = os.path.join(log_file_path, current_log_file_name)
    if os.path.exists(complete_path):
        new_file = False
    else:
        new_file = True

    with open(complete_path, "a+", newline="\n", encoding="utf-8") as fd:
        fieldnames = ["Ignore","TimeGenerated","EventID","EventCategory","RecordNumber", "StringInserts"]
        writer = csv.DictWriter(fd, fieldnames=fieldnames)
        if new_file:
            writer.writeheader()

        for event_index in range(event_index, len(events_of_interest)):
            event = events_of_interest[event_index]
            print(last_written, event.TimeGenerated)
            # check if the event is within the same month as the last written one,
            # otherwise change the file
            if event.TimeGenerated.month == last_written.month:
                last_written = event.TimeGenerated
                writer.writerow({'Ignore': 'N',
                                 'TimeGenerated': event.TimeGenerated,
                                 'EventID': event.EventID,
                                 'EventCategory': event.EventCategory,
                                 'RecordNumber': event.RecordNumber,
                                 'StringInserts': event.StringInserts})
            else:
                break

        # check if the last event was written to file, or wheter a break
        # occured to switch files
        if event and last_written != event.TimeGenerated:
            last_written = event.TimeGenerated
        else:
            break

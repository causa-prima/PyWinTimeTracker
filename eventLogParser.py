import win32evtlog
from datetime import datetime
from collections import deque
from time import mktime
import csv
import os


def isSameDay(date1, date2):
    return date1.day == date2.day and date1.month == date2.month and date1.year == date2.year

log_file_path = r"C:\Users\Sebastian Kieritz\Documents\WorkLog"
log_file_name = r"EventList-{}-{:0>2}.csv"
# default date for last run
last_run = None

DO_NOT_IGNORE = 'N'
EVENT_LOCK = 4800
EVENT_UNLOCK = 4801

# get the date and time of the last run
for file_name in sorted(os.listdir(log_file_path), reverse=True):
    if file_name.startswith("EventList-") and file_name.endswith(".csv"):
        with open(os.path.join(log_file_path, file_name)) as csv_file:
            csv_deque = deque(csv.DictReader(csv_file), 1)
            if csv_deque:
                last_run = datetime.strptime(csv_deque[0]["TimeGenerated"], "%Y-%m-%d %H:%M:%S")
                break

print('last run:', last_run or 'Never')
last_run = last_run or datetime(2017, 1, 1, 0, 0, 0)

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

    if last_event and not isSameDay(event.TimeGenerated, last_event.TimeGenerated):
        # the day changed, append the current and perhaps the old event
        if not last_event_was_appended:
            events_of_interest.append(last_event)
        events_of_interest.append(event)
        last_event_was_appended = True
    elif event.EventID in (EVENT_LOCK, EVENT_UNLOCK):
        events_of_interest.append(event)
        last_event_was_appended = True
    else:
        last_event_was_appended = False
    last_event = event

# save last event if it was not yet appended and the day differs from
# the day of the last run (hence the last run must have been the last event of that day)
if last_event and not last_event_was_appended and not isSameDay(last_event.TimeGenerated, last_run):
    events_of_interest.append(last_event)

# write all events of interest to the corresponding csv file
print('Adding {} new events to the event logs'.format(len(events_of_interest)))
event_index = 0
events_of_interest.reverse()
while event_index < len(events_of_interest):
    current_write_year = events_of_interest[event_index].TimeGenerated.year
    current_write_month = events_of_interest[event_index].TimeGenerated.month
    current_log_file_name = log_file_name.format(current_write_year, current_write_month)    
    complete_log_file_path = os.path.join(log_file_path, current_log_file_name)    
    is_new_file = not os.path.exists(complete_log_file_path)

    with open(complete_log_file_path, "a+", newline="\n", encoding="utf-8") as fd:
        fieldnames = ["Ignore","TimeGenerated","EventID"]
        writer = csv.DictWriter(fd, fieldnames=fieldnames)
        if is_new_file:
            writer.writeheader()

        for event_index in range(event_index, len(events_of_interest)):
            event = events_of_interest[event_index]
            # check if the event is within the same month as the last written one,
            # otherwise break and change the file
            if event.TimeGenerated.month == current_write_month:
                print(event.TimeGenerated, event.EventID)
                writer.writerow({'Ignore': DO_NOT_IGNORE,
                                 'TimeGenerated': event.TimeGenerated,
                                 'EventID': event.EventID})
            else:
                break

        # check if the last event was written to file, or whether a break
        # occurred to switch files
        if not event or event.TimeGenerated.month == current_write_month:
            break

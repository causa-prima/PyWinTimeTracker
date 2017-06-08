import csv
import os
from datetime import date, datetime, time, timedelta


log_file_path = r"C:\Users\SebastianSchultz\Documents\WorkLog"

class LockPeriod(object):
    def __init__(self, begin = None):
        self.begin = begin
        self.end = None

    def __repr__(self):
        return "<{} {}-{}>".format(self.__class__.__name__, self.begin, self.end)

    @property
    def duration(self):
        if self.begin and self.end:
            return self.end - self.begin
        else:
            return timedelta()

class Workday(object):
    def __init__(self, begin = None):
        self.begin = begin
        self.end = None
        self.lock_periods = []
        self.__lunch_duration = None

    def __repr__(self):
        out_string = "<{} {} {}-{} LL: {} WH: {} CWH: {} OT: {:>8}>"
        return out_string.format(self.__class__.__name__,
                                 self.begin.date(),
                                 self.begin.time(),
                                 self.end.time(),
                                 self.lunch_duration,
                                 self.working_hours,
                                 self.corrected_working_hours,
                                 str(self.overtime) if self.overtime > timedelta(0) else "-{}".format(-self.overtime))

    def __compute_lunch_duration(self):
        self.__lunch_duration = None
        for lock_period in self.lock_periods:
            if time(10,30) < lock_period.begin.time() < lock_period.end.time() < time(13,30):                
                if self.__lunch_duration is None or self.__lunch_duration < lock_period.duration:
                    self.__lunch_duration = lock_period.duration
                
        if self.__lunch_duration is None:
            self.__lunch_duration = timedelta(minutes=30)

    @property
    def working_hours(self):
        if self.begin and self.end:
            return self.end - self.begin
        else:
            return timedelta()

    @property
    def corrected_working_hours(self):
        return self.working_hours - self.lunch_duration

    @property
    def lunch_duration(self):
        if self.__lunch_duration is None:
            self.__compute_lunch_duration()
        return self.__lunch_duration

    @property
    def overtime(self):
        return self.corrected_working_hours - timedelta(hours=8)
    

workdays = []
current_workday = Workday()

for file_name in sorted(os.listdir(log_file_path)):
    if file_name.startswith("EventList-") and file_name.endswith(".csv"):
        print("processing", file_name)
        with open(os.path.join(log_file_path, file_name)) as csv_fd:
            reader = csv.reader(csv_fd)
            next(reader, None) # skip the header
            for row in reader:
                time_generated = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S")        
                if current_workday.begin is None or current_workday.begin.date() != time_generated.date():
                    if current_workday.begin is not None:
                        workdays.append(current_workday)
                    current_workday = Workday(begin = time_generated)
                current_workday.end = time_generated

                # 4800 = lock, 4801 = unlock
                event_id = row[2]
                if event_id == "4800":
                    current_workday.lock_periods.append(LockPeriod(time_generated))
                elif event_id == "4801":
                    last_lock_period = current_workday.lock_periods[-1]
                    if last_lock_period and last_lock_period.end is None:
                        last_lock_period.end = time_generated
                    else:
                        print("LockPeriod ended on {} without beginning!".format(time_generated))
                        
            # append the last workday if it is a complete workday (handling the last line for completed months)
            if current_workday.end.date() == date.today():
                # If the last processed workday is today, it's end is not yet reached, but will be some time in
                # the future. Hence, set it to the current time.
                current_workday.end = datetime.now()
            """if current_workday.end != current_workday.begin and workdays[-1] is not current_workday:
                print('appending')
                workdays.append(current_workday)"""
        
print()
overtime = timedelta()
for workday in workdays:
    print(workday)
    overtime += workday.overtime

print()
print("aggregated overtime: {}".format(overtime))

import csv
import os
from datetime import date, datetime, time, timedelta


log_file_path = r"C:\Users\SebastianSchultz\Documents\WorkLog"

class LockPeriod(object):
    def __init__(self, begin = None):
        self.begin = begin
        self.end = None

    def __repr__(self):
        return "<{} {} {}-{}>".format(self.__class__.__name__, self.begin.date(), self.begin.time(), self.end.time())

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
        self.__lunch_period = None

    def __repr__(self):
        out_string = "<{} {} {}-{} LL: {} WH: {!s:>8} CWH: {!s:>8} OT: {:>8}>"
        return out_string.format(self.__class__.__name__,
                                 self.begin.date(),
                                 self.begin.time(),
                                 self.end.time(),
                                 self.lunch_period,
                                 self.working_hours,
                                 self.corrected_working_hours,
                                 str(self.overtime) if self.overtime > timedelta(0) else "-{}".format(-self.overtime))

    def __compute_lunch_period(self):
        self.__lunch_period = None
        for lock_period in self.lock_periods:
            if time(10,30) < lock_period.begin.time() < lock_period.end.time() < time(13,30):
                if self.__lunch_period is None or self.__lunch_period.duration < lock_period.duration:
                    self.__lunch_period = lock_period
                
        if self.__lunch_period is None:
            lunch_period_helper = datetime(self.begin.year, self.begin.month, self.begin.day, hour=12)
            lunch_period = LockPeriod(begin=lunch_period_helper)
            lunch_period.end = lunch_period_helper + timedelta(minutes=30)
            self.__lunch_period = lunch_period

    @property
    def working_hours(self):
        if self.begin and self.end:
            return self.end - self.begin
        else:
            return timedelta()

    @property
    def corrected_working_hours(self):
        return self.working_hours - self.lunch_period.duration

    @property
    def lunch_period(self):
        if self.__lunch_period is None:
            self.__compute_lunch_period()
        return self.__lunch_period

    @property
    def overtime(self):
        return self.corrected_working_hours - timedelta(hours=8)
    

workdays = []
current_workday = Workday()

for file_name in sorted(os.listdir(log_file_path)):
    if file_name.startswith("EventList-") and file_name.endswith(".csv"):
        print("processing", file_name)
        with open(os.path.join(log_file_path, file_name)) as csv_fd:
            reader = csv.DictReader(csv_fd)
            #next(reader, None) # skip the header
            for row in reader:
                if row['Ignore'] == 'Y':
                    continue
                time_generated = datetime.strptime(row['TimeGenerated'], "%Y-%m-%d %H:%M:%S")
                if current_workday.begin is None or current_workday.begin.date() != time_generated.date():
                    if current_workday.begin is not None:
                        workdays.append(current_workday)
                    current_workday = Workday(begin = time_generated)
                current_workday.end = time_generated

                # 4800 = lock, 4801 = unlock
                event_id = row['EventID']
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
        
print()
overtime = timedelta()
for workday in workdays:
    print(workday)
    overtime += workday.overtime

print()
print("aggregated overtime: {} ({:.2f} hours)".format(overtime, overtime.total_seconds()/3600))

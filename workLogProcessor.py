import csv
import os
from datetime import date, datetime, time, timedelta


log_file_path = r"C:\Users\Sebastian Kieritz\Documents\WorkLog"
DAILY_WORKING_HOURS = 8

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

    @property
    def is_lunch_period(self):
        if time(10, 30) < self.begin.time() < self.end.time() < time(13, 30):
            return True
        else:
            return False


class Workday(object):
    def __init__(self, begin = None):
        self.begin = begin
        self.end = None
        self.lock_periods = []
        self.__lunch_period = None

    def __repr__(self):
        out_string = "<{} {} {}-{}>"
        return out_string.format(self.__class__.__name__,
                                 self.begin.date(),
                                 self.begin.time(),
                                 self.end.time())

    def __compute_lunch_period(self):
        self.__lunch_period = None
        for lock_period in self.lock_periods:
            if lock_period.is_lunch_period:
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
        return self.corrected_working_hours - timedelta(hours=DAILY_WORKING_HOURS)


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
                    current_lock_period = LockPeriod(time_generated)
                elif event_id == "4801":
                    if current_lock_period and current_lock_period.end is None:
                        current_lock_period.end = time_generated
                        current_workday.lock_periods.append(current_lock_period)
                        current_lock_period = None
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
    out_string = "{} LL: {} ({}-{}) WH: {!s:>8} CWH: {!s:>8} OT: {:>8}"
    print(out_string.format(workday, workday.lunch_period.duration, workday.lunch_period.begin.time(), workday.lunch_period.end.time(), workday.working_hours, workday.corrected_working_hours, str(workday.overtime) if workday.overtime > timedelta(0) else "-{}".format(-workday.overtime)))
    overtime += workday.overtime

print()
SECONDS_IN_HOUR = 3600
overtime_in_hours = overtime.total_seconds()/SECONDS_IN_HOUR
overtime_in_working_days = overtime.total_seconds()/SECONDS_IN_HOUR/DAILY_WORKING_HOURS
print("aggregated overtime: {} ({:.2f} hours / {:.2f} working days)".format(overtime, overtime_in_hours, overtime_in_working_days))

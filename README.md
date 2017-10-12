# PyWinTimeTracker
Uses the windows event log to track the times the computer was active.

Prerequisites:

* Python3
* pypiwin32 (pip install pypiwin32)
* enabled logging of lock- and unlock-events (optional)
    1. Open the group policy editor
    2. Navigate to "Windows Settings > Security Settings > Advanced Audit Policy Configuration > System Audit Policies - Local Group > Logon/Logoff"
    3. Right click Audit Logoff and Audit Logon
    4. Click "Properties"
    5. Enable "Configure the following audit events:"
    6. Enable "Success"
* eventLogParser.py needs to be run with administrator privileges to be able to access the event log.

## Plans for further improvements

* introduce a config-file for all configurable settings (e.g. file paths, default lunch length, etc.)
* ability to add corrections (e.g. paid overtimes)
* increase foolproofness
  * check existence of paths
  * handle insufficient rights
  * ?
* document usage

## Automated parsing of the event log

Automatically parsing the event log for the wanted events by executing eventLogParser.py is possible, e.g. using a scheduled task that is executed periodically, e.g. on log on:

1. Open the Task Scheduler.
2. Click "Create Basic Task...", fill the name and description, and click "Next" to advance to the "Trigger" step.
3. Choose the when the task should be executed, e.g. "When I log on" and click "Next" to advance to the "Action" step.
4. Choose "Start a program" and click "Next" to advance to the "Start a Program" step.
5. In the "Program/scipt:" field, insert the path to you python executable, e.g. "C:\Program Files (x86)\Python36-32\python.exe"
6. In the "Add arguments (optional)" field, insert the path to eventLogParser.py, e.g. C:\PyWinTimeTracker\eventLogParser.py and click "Next" to advance to the "Finish" step.
7. Check "Open the Properties dialog for this task when I click Finish" and click "Finish" to open the properties dialog.
8. On the "General" page, check "Run with highest privileges". (TODO: Is this needed?).
9. Optional: On the "Conditions" page, uncheck "Start the task only if the computer is on AC power".
10. Optional: On the "Settings" page, check "Run task as soon as popssible after a scheduled start is missed" (TODO: Does it make sense to enable this option?)
11. Click "OK" to close the dialog and save the task.

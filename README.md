# PyWinTimeTracker
Uses the windows event log to track the times the computer was active.

Prerequisites:

* Python3
* pypiwin32 (pip install pypiwin32)
* enabled logging of lock- and unlock-events (optional)
* eventLogParser.py needs to be run with administrator privileges to be able to access the event log.

## Plans for further improvements

* introduce a config-file for all configurable settings (e.g. file paths, default lunch length, etc.)
* ability to add corrections (e.g. paid overtimes)
* increase foolproofness
  * check existence of paths
  * handle insufficient rights
  * ?
* document usage

Timetracker : Excel extract with times grouped by parent and by team member
===================

This tool can export times over a time period, recursively grouped by parents, and grouped by team member

This is a fork from https://github.com/7pace/timetracker-api-samplecode , which is a code sample for the Timetracker API.

## TimetrackerOdataClient usage

Command line parameters:

VSTS usage (token auth): 
```
TimetrackerExcelExporter.exe TIMETRACKER_SERVICE_URI -t TOKEN -f VSTS_ACCOUNT_URL -v VSTS_TOKEN 
```


On-premise usage (NTLM auth):
```
TimetrackerExcelExporter.exe TIMETRACKER_SERVICE_URI -w -f TFS_URL_WITH_COLLECTION
```
## Parameters

|   | TFS  | VSTS  |
|---|---|---|
| TIMETRACKER_SERVICE_URI  | [timetrackerServiceUrl:Port]/api/[CollectionName]/odata  |  https://[accountName].timehub.7pace.com/api/odata |
|-f| TFS URL (like http://tfs:8080/tfs)|VSTS Account URL (https://[accountName].visualstudio.com)|
| -t  | -  | Timetracker API Token  |
| -v  | -  | VSTS Personal token.  |
| -w  | no value, tells application to use Windows Credentials  | -  |
| --from  | Start date : get time tracked this date or after   | idem |
| --to  |  End date : get time tracked this date or before | idem |



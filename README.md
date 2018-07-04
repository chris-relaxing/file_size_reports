# File Size Reports - Readme

The File Size Report script allows you to compare product file sizes quarter over quarter. This tool creates a QoQ report that identifies drops and increases in file sizes by comparing byte counts between files that have the same name in both quarters. The idea is to compare a product folder for the current quarter with the folder that contains the same product from last quarter.<br><br>

### File Size Report script inputs:<br>
For inputs, you provide the path for the previous quarter product, the path for the new quarter product, and type a name for your report and the script will compare file sizes for both quarters and create a report in the “reports” folder in the same directory where the script lives.
![alt text](http://bluegalaxy.info/images/file-size-reports-input.png)

<br><br>
Click “Create Report” when selections have been made. The script will alert you when the report has been written.
![alt text](http://bluegalaxy.info/images/file-size-reports-alert.png)<br><br>

If there is not a match in file names between quarters, then the file will be listed under “Unmatched File Names” with no comparison made. Example of output:
![alt text](http://bluegalaxy.info/images/file-size-reports-output.png)<br><br>

### Use Cases:<br>
This script came about due to the fact that P&RP often asks after product releases why there is drop in file sizes between the the current quarter’s product release folder and the previous quarter’s product release folder. This script addresses the question by creating a report that details exactly which files changed in file size between quarters.<br><br>
This script can also be useful as a product validation tool to ensure that there are no unexpected file size drops.

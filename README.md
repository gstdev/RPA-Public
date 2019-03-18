# Use Case
Client needed hundreds of Word documents (.doc and .docx) converted to PDF format.  There are existing tools to accomplish batch PDF conversions but the problem is that a majority of these documents contained pending changes and those revisions/comments were displayed in the PDF.  We did not find a tool that contained configuration settings to approve all changes prior to the PDF conversion.

# Example of converted PDF with pending document changes
![Image of Pending Changes](https://s3.amazonaws.com/gst-public-share/github/gst_revisions_example.png)

# Manual process to approve changes, change review display format, and export document to PDF
![Gif of Manual Process](https://media.giphy.com/media/1ynCRohvQBcBimh1VB/giphy.gif)

# How can we utilize technology and avoid many hours of tedious work?  Let's see if this is a good fit for Robotic Process Automation (RPA)
We have been interested in RPA for over a year now and spent some time with UiPath but didn't have an actual use case to do a proof of concept.  Currently the three RPA marketplace leaders are <a href="https://www.uipath.com" target="_blank">UiPath</a>, <a href="https://www.automationanywhere.com/" target="_blank">Automation Anywhere</a>, and <a href="https://www.blueprism.com" target="_blank">Blue Prism</a>.  We decided to use UiPath since they offer a free <a href="https://www.uipath.com/developers/community-edition-download" target="_blank">Community Edition</a>.

# Project Setup
* Modify Word ribbon (add option to Export as PDF to review tab), hotkey sequence may need to be altered in project based on your Word ribbon
* Ensure that open file checkbox is not checked in dialog box during Export as PDF process (this check/uncheck is persistent so you want to ensure it is unchecked prior to running project)
* Install UiPath Studio
* Clone git repository
* Open up UiPath Studio
* Add/Update Dependencies (Mail, Word, Excel, UIAutomation, System)
![Manage Packages](https://s3.amazonaws.com/gst-public-share/github/manage_packages.png)
* Open project (ApproveConvertRename/project.json)
* Double click on ApproveConvertRename.xaml to view the project
![View Project](https://s3.amazonaws.com/gst-public-share/github/view_project.png)
* Double click on any elements to drill into nested sequences
* Click on Activities tab (in left side navigation) to view all of the available elements
* Click on Variables tab (in main project) to view all of the variables (can only see variables based on the defined scoping)
![Variables](https://s3.amazonaws.com/gst-public-share/github/variable_configuration.png)
* Create necessary directories and update variables in order to succesfully run this sample project
| Variable  | Scope |
| ------------- | ------------- |
| initialPath  | Approve, Convert, and Rename  |
| approvedWordDirectory  | Approve, Convert, and Rename  |
| convertedPDFDirectory  | Approve, Convert, and Rename  |
| errorLog  | Try Catch  |
| excelFile  | Rename Files  |
| renamedFileDirectory  | Cell Match  |
* Verify all of your hotkeys are the same
* Since the main workflow is "Approve, Convert, and Rename" and it exists in both the Try Catch and the Finally Retry Scope, I would recommend making necessary modifications in just the Try Catch and then copy/paste into the Finally Retry Scope.  I have this in place since I would occassionally receive unexpected exceptions.

# Use Case
Client needed hundreds of Word documents (.doc and .docx) converted to PDF format.  There are existing tools to accomplish batch PDF conversions but the problem is that a majority of these documents contained pending changes and those revisions/comments were displayed in the PDF.  We did not find a conversion tool that had configuration settings to approve all changes prior to the PDF conversion.

# Example of converted PDF with pending document changes
![Image of Pending Changes](https://s3.amazonaws.com/gst-public-share/github/gst_revisions_example.png)

# Manual process to approve changes, change review display format, and export document to PDF
![Gif of Manual Process](https://media.giphy.com/media/1ynCRohvQBcBimh1VB/giphy.gif)

# How can we utilize technology and avoid many hours of tedious work?  Let's see if this is a good fit for Robotic Process Automation (RPA)
We have been interested in RPA since early 2018 and spent some time with UiPath but didn't have an actual use case to do a proof of concept.  Currently the three RPA marketplace leaders are <a href="https://www.uipath.com" target="_blank">UiPath</a>, <a href="https://www.automationanywhere.com/" target="_blank">Automation Anywhere</a>, and <a href="https://www.blueprism.com" target="_blank">Blue Prism</a>.  We decided to use UiPath since they offer a free <a href="https://www.uipath.com/developers/community-edition-download" target="_blank">Community Edition</a>.

# Summary
This project took roughly 9 hours to put together.  Manually executing this task for one document would roughly take 60 seconds to execute and there is risk when it comes to moving the files or matching/renaming the file.  Each document took about 24 seconds to process using RPA (this includes 10 seconds of built in delays) without any errors.  We now have an accurate approach to execute this task for an unlimited number of documents at 2.5 times the speed of a person.

Imagine if this task was one person's full time job and they were getting paid the federal minimum wage of $7.25.  That would mean they could process about 480 documents for $58 per day (this doesn't even take into account the cost of errors).  Now imagine if you had a team of x people performing this same task!

| Process  | Documents | Time | Cost |
| ------------- | ------------- | ------------- | ------------- |
| 1 Person  | 480  | 8 hours  | $58  |
| 5 People  | 2,400  | 8 hours  | $290  |
| 10 People  | 4,800  | 8 hours  | $580  |
| 50 People  | 24,000  | 8 hours  | $2,900  |
| 1 Bot  | 1,200  | 8 hours  | N/A  |
| 1 Bot  | 3,600  | 24 hours  | N/A  |
| 5 Bots  | 18,000  | 24 hours  | N/A  |
| 10 Bots  | 36,000  | 24 hours  | N/A  |

You can easily see the potential benefits of RPA for enterprises that have a lot of manual tasks being performed.  Here is an <a href="https://www.uipath.com/blog/rpa-and-the-roi-conundrum" target="_blank">interesting article</a> that goes over RPA and the ROI conundrum.  Feel free to send us an email at rpa@gstdev.com if you're interested in learning more about RPA.

# Project Setup
* Modify Word ribbon (add option to Export as PDF to review tab)
![Word Ribbon](https://s3.amazonaws.com/gst-public-share/github/word_ribbon.png)
* Ensure that open file checkbox is not checked in dialog box during Export as PDF process (this check/uncheck is persistent so you want to ensure it is unchecked prior to running project)
![Publish Checkbox](https://s3.amazonaws.com/gst-public-share/github/open_file_checkbox.png)
* Install UiPath Studio
* Clone git repository
* Open up UiPath Studio
* Add/Update Dependencies (Mail, Word, Excel, UIAutomation, System)
![Manage Packages](https://s3.amazonaws.com/gst-public-share/github/manage_packages.png)
* Open project (ApproveConvertRename/project.json)
* Double click on ApproveConvertRename.xaml (in left side navigation) to view the project
![View Project](https://s3.amazonaws.com/gst-public-share/github/view_project.png)
* Double click on any elements to drill into nested sequences
* Click on Activities tab (in left side navigation) to view all of the available actions
* Click on Variables tab (in center window - workflow designer) to view all of the variables (can only see variables based on the scoping)
![Variables](https://s3.amazonaws.com/gst-public-share/github/variable_configuration.png)
* Create necessary directories and update variables in order to successfully run this sample project

| Variable  | Scope | Note |
| ------------- | ------------- | ------------- |
| initialPath  | Approve, Convert, and Rename  | Default Document Path  |
| approvedWordDirectory  | Approve, Convert, and Rename  | Default Approved Document Path  |
| convertedPDFDirectory  | Approve, Convert, and Rename  | Default Converted Document Path  |
| errorLog  | Try Catch  | Error Log Text File  |
| excelFile  | Rename Files  | Excel Spreadsheet for File Rename  |
| renamedFileDirectory  | Cell Match  | Default Renamed Document Path  |

* Verify all of your Word hotkeys match (set in Root Directory Loop scope).  If yours are different, then you will need to modify the Root Directory Loop sequences so that they match your hotkeys.
  * Approve all comments
    * Alt + r
    * a2l
  * Set markup drop down to "No Markup" selection, just execute what's inside the parentheses of [k()] (i.e. down arrow or enter key)
    * Alt + r
    * td[k(down)][k(down)][k(enter)]
  * Export as PDF
    * Alt + r
    * y
    * Enter
  * Save document
    * Ctrl + s
  * Close application
    * Alt + F4
  
* Since the main workflow is "Approve, Convert, and Rename" and it exists in both the Try Catch and the Finally Retry Scope, I would recommend making necessary modifications in just the Try Catch and then copy/paste into the Finally Retry Scope.  I have this in place since I would occasionally receive unexpected exceptions.

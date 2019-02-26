# Use Case
Client needed hundreds of Word documents (.doc and .docx) converted to PDF format.  There are existing tools to accomplish batch PDF conversions but the problem is that a majority of these documents contained pending changes and those revisions/comments were displayed in the PDF.  We did not find a tool that contained configuration settings to approve all changes prior to the PDF conversion.

# Example of converted PDF with pending document changes
![Image of Pending Changes](https://s3.amazonaws.com/gst-public-share/github/gst_revisions_example.png)

# Manual process to approve changes, change review display format, and export document to PDF
![Gif of Manual Process](https://media.giphy.com/media/1ynCRohvQBcBimh1VB/giphy.gif)

# How can we utilize technology and avoid many hours of tedious work?  Let's see if this is a good fit for Robotic Process Automation (RPA)
We have been interested in RPA for over a year now and spent some time with UiPath but didn't have an actual use case to do a proof of concept.  Currently the three RPA marketplace leaders are <a href="https://www.uipath.com" target="_blank">UiPath</a>, <a href="https://www.automationanywhere.com/" target="_blank">Automation Anywhere</a>, and <a href="https://www.blueprism.com" target="_blank">Blue Prism</a>.  We decided to use UiPath since they offer a free <a href="https://www.uipath.com/developers/community-edition-download" target="_blank">Community Edition</a>.

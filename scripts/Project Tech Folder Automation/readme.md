# Overview

This set of scripts is used by a small company to automate some project process flows. They track project information in a google spreadsheet and then use google drive to store information for each project.

This script watches for a specific column to change to a specific value - a project status in this case - and then fires off the second part that recursively copies the template folder to the new folder location.

# Usage

## Simple
- Copy contents to local apps script project of the spreadsheet to watch
- Replace config.get() instances with your specific drive folder ids
- Add managed trigger to fire onTriggeredUpdate from the Spreadsheet On Edit event

## Versioned Library

### Setup
- Create a Google Apps Script project with the contents of this project.
- Create a new GCP Project
- Add Spreadsheet API & Drive API to your GCP Project
- Configure the Apps Script Project to use your GCP Project
- Create a new Apps Script deployment of type Library
- New revisions can be saved/created by creating a new Library deployment

### Usage
- In the spreadsheet to watch, add as a library using the deployed script id
- Add a function that is triggered that calls `ProjectTechFolderAutomation.onTriggeredEvent(e)`.
- Add managed trigger to fire your new function from the Spreadsheet On Edit event
- On future library updates, you can change the revision used without needing to re-add the library

# Future Thoughts
- make config.get() calls work instead of needing replaced
- tokenize file substitution patters
- split out into more abstracted scripts: 
    - utils library
    - spreadsheet watcher
    - recursive copy of files & directories
    - recursive update of files
- make more generic for wider use, even outside of ESI
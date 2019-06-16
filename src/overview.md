# WorkItem Updater

## Overview
The WorkItem Updater task can update workitems using a Query or that are linked to the Build/Release.  
The following workitem fields can be updated:
- Update the state for workitems.  
- Update the assignee for workitems.  
- Update the swimlane or board-column for workitems.  
- Add the build as Development Link to the workitems.
  
By adding this task to specific milestones in a build/release pipeline, you can create an automated kanban board.  
An example would be to set the state to 'Resolved' as the last step of a build and to 'Deployed' as the last step of a release.  
With this task the state of workitems is always reflecting reality and developers don't need to manually update workitems anymore.  
  
A preview of what the task can do, can be seen in this recording:  
  
![AutoKanban](img/AutoKanban.gif)  
  
## Settings
![settings](img/Settings.png)  
  
## Version History
### 2.5.791
- Functionality to add custom list of Fields to update 
- Functionality to allow to bypass workitem type rules 
- Functionality to allow to fail task when there's no workitem to update 
### 2.4.784
- Functionality to add tags to a workitem
- Functionality to remove existing tags from a workitem  
### 2.2.764
- Functionality to, while deploying a release to an environment, update all workitems since the previously completed deployment.
### 2.1.762
- Be able to update multiple workitem types at once.
### 2.0.23
- Rebuild extension using Node.js.
- Added option to specify the source of the workitems. This can be a Query or workitems linked to the build.
### 1.5.9
- Update 'Assigned To' with a fixed user.
- Update 'Assigned To' with option to unassign the workitem.
### 1.4.16
- Filter the workitems to update by specifying the current state a workitem needs to have.
### 1.3.6
- Update Board-Swimlane.
- Update Board-Column.
### 1.2.0
- Add Build as Development link.
- Update 'Assigned To' with requester of the build.
### 1.1.8
- Initial Version.

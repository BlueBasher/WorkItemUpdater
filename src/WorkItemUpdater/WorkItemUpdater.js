"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const tl = require("vsts-task-lib/task");
const settings_1 = require("./settings");
const WebApi_1 = require("vso-node-api/WebApi");
const VSSInterfaces_1 = require("vso-node-api/interfaces/common/VSSInterfaces");
const WorkItemTrackingInterfaces_1 = require("vso-node-api/interfaces/WorkItemTrackingInterfaces");
function main() {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            let vstsWebApi = getVstsWebApi();
            let settings = getSettings();
            let workItemTrackingClient = yield vstsWebApi.getWorkItemTrackingApi();
            let workItemRefs = yield getWorkItemsRefs(vstsWebApi, workItemTrackingClient, settings);
            if (!workItemRefs || workItemRefs.length === 0) {
                console.log("No workitems found to update.");
            }
            else {
                tl.debug("Loop workItemsRefs");
                yield asyncForEach(workItemRefs, (workItemRef) => __awaiter(this, void 0, void 0, function* () {
                    tl.debug("Found WorkItemRef: " + workItemRef.id);
                    let workItem = yield workItemTrackingClient.getWorkItem(parseInt(workItemRef.id), null, null, WorkItemTrackingInterfaces_1.WorkItemExpand.Relations);
                    console.log("Found WorkItem: " + workItem.id);
                    switch (settings.updateAssignedToWith) {
                        case "Creator": {
                            let creator = workItem.fields["System.CreatedBy"];
                            tl.debug("Using workitem creator user '" + creator + "' as assignedTo.");
                            settings.assignedTo = creator;
                            break;
                        }
                        case "FixedUser": {
                            tl.debug("Using fixed user '" + settings.assignedTo + "' as assignedTo.");
                            break;
                        }
                        case "Unassigned": {
                            tl.debug("Using Unassigned as assignedTo.");
                            settings.assignedTo = null;
                            break;
                        }
                        default: {
                            tl.debug("Setting assignedTo to requester for build '" + settings.requestedFor + "'.");
                            settings.assignedTo = settings.requestedFor;
                            break;
                        }
                    }
                    yield updateWorkItem(workItemTrackingClient, workItem, settings);
                }));
                tl.debug("Finished loop workItemsRefs");
            }
            tl.setResult(tl.TaskResult.Succeeded, "");
        }
        catch (error) {
            tl.debug("Caught an error in main: " + error);
            tl.setResult(tl.TaskResult.Failed, error);
        }
    });
}
function asyncForEach(array, callback) {
    return __awaiter(this, void 0, void 0, function* () {
        for (let index = 0; index < array.length; index++) {
            yield callback(array[index], index, array);
        }
    });
}
function getVstsWebApi() {
    let endpointUrl = tl.getVariable("System.TeamFoundationCollectionUri");
    let accessToken = tl.getEndpointAuthorizationParameter('SYSTEMVSSCONNECTION', 'AccessToken', false);
    let credentialHandler = WebApi_1.getHandlerFromToken(accessToken);
    let webApi = new WebApi_1.WebApi(endpointUrl, credentialHandler);
    return webApi;
}
function getSettings() {
    let settings = new settings_1.Settings();
    settings.buildId = parseInt(tl.getVariable("Build.BuildId"));
    settings.projectId = tl.getVariable("System.TeamProjectId");
    settings.requestedFor = tl.getVariable("Build.RequestedFor");
    settings.workitemsSource = tl.getInput("workitemsSource");
    settings.workitemsSourceQuery = tl.getInput("workitemsSourceQuery");
    settings.workItemType = tl.getInput("workItemType");
    settings.workItemState = tl.getInput("workItemState");
    settings.workItemCurrentState = tl.getInput("workItemCurrentState");
    settings.workItemKanbanLane = tl.getInput("workItemKanbanLane");
    settings.workItemKanbanState = tl.getInput("workItemKanbanState");
    settings.workItemDone = tl.getBoolInput("workItemDone");
    settings.linkBuild = tl.getBoolInput("linkBuild");
    settings.updateAssignedTo = tl.getInput("updateAssignedTo");
    settings.updateAssignedToWith = tl.getInput("updateAssignedToWith");
    settings.assignedTo = tl.getInput("assignedTo");
    tl.debug("BuildId " + settings.buildId);
    tl.debug("ProjectId " + settings.projectId);
    tl.debug("requestedFor " + settings.requestedFor);
    tl.debug("workitemsSource " + settings.workitemsSource);
    tl.debug("workitemsSourceQuery " + settings.workitemsSourceQuery);
    tl.debug("workItemType " + settings.workItemType);
    tl.debug("WorkItemState " + settings.workItemState);
    tl.debug("workItemCurrentState " + settings.workItemCurrentState);
    tl.debug("updateWorkItemKanbanLane " + settings.workItemKanbanLane);
    tl.debug("WorkItemKanbanState " + settings.workItemKanbanState);
    tl.debug("WorkItemDone " + settings.workItemDone);
    tl.debug("updateAssignedTo " + settings.updateAssignedTo);
    tl.debug("updateAssignedToWith " + settings.updateAssignedToWith);
    tl.debug("assignedTo " + settings.assignedTo);
    return settings;
}
function getWorkItemsRefs(vstsWebApi, workItemTrackingClient, settings) {
    return __awaiter(this, void 0, void 0, function* () {
        if (settings.workitemsSource === 'Build') {
            console.log("Using Build as WorkItem Source");
            let buildClient = yield vstsWebApi.getBuildApi();
            let workItemRefs = yield buildClient.getBuildWorkItemsRefs(settings.projectId, settings.buildId);
            return workItemRefs;
        }
        else if (settings.workitemsSource === 'Query') {
            console.log("Using Query as WorkItem Source");
            var queryResult = yield workItemTrackingClient.queryById(settings.workitemsSourceQuery, {
                project: null,
                projectId: settings.projectId,
                team: null,
                teamId: null
            });
            var result = [];
            queryResult.workItems.forEach((workItem) => {
                result.push({
                    id: workItem.id.toString(),
                    url: workItem.url
                });
            });
            return result;
        }
        return null;
    });
}
function updateWorkItem(workItemTrackingClient, workItem, settings) {
    return __awaiter(this, void 0, void 0, function* () {
        tl.debug("Updating  WorkItem: " + workItem.id);
        if (settings.workItemType.split(',').indexOf(workItem.fields["System.WorkItemType"]) >= 0) {
            if (settings.workItemCurrentState && settings.workItemCurrentState !== "" && settings.workItemCurrentState.split(',').indexOf(workItem.fields["System.State"]) === -1) {
                console.log("Skipped WorkItem: " + workItem.id + " State: '" + workItem.fields["System.State"] + "' => Only updating if state in '" + settings.workItemCurrentState) + "'";
                return;
            }
            console.log("Updating WorkItem " + workItem.id);
            let document = [];
            let kanbanLane = getWorkItemFields(workItem, (f) => f.endsWith("Kanban.Lane"));
            tl.debug("Found KanbanLane: " + kanbanLane);
            let kanbanColumn = getWorkItemFields(workItem, (f) => f.endsWith("Kanban.Column"));
            tl.debug("Found KanbanColumn: " + kanbanColumn);
            let kanbanColumnDone = getWorkItemFields(workItem, (f) => f.endsWith("Kanban.Column.Done"));
            tl.debug("Found KanbanColumnDone: " + kanbanColumnDone);
            if (settings.workItemState && settings.workItemState !== "") {
                addPatchOperation("/fields/System.State", settings.workItemState, document);
            }
            if (settings.workItemKanbanLane && settings.workItemKanbanLane !== "" && kanbanLane.length > 0) {
                kanbanLane.forEach((lane, index) => {
                    addPatchOperation("/fields/" + lane, settings.workItemKanbanLane, document);
                });
            }
            if (settings.workItemKanbanState && settings.workItemKanbanState !== "" && kanbanColumn.length > 0) {
                kanbanColumn.forEach((column, index) => {
                    addPatchOperation("/fields/" + column, settings.workItemKanbanState, document);
                });
            }
            if (kanbanColumnDone.length > 0) {
                kanbanColumnDone.forEach((columnDone, index) => {
                    addPatchOperation("/fields/" + columnDone, settings.workItemDone, document);
                });
            }
            if (settings.linkBuild) {
                let buildRelationUrl = "vstfs:///Build/Build/$buildId";
                let buildRelation = workItem.relations.find(r => r.url === buildRelationUrl);
                if (buildRelation === null) {
                    console.log("Linking Build " + settings.buildId + " to WorkItem " + workItem.id);
                    let relation = {
                        rel: "ArtifactLink",
                        url: buildRelationUrl,
                        attributes: {
                            name: "Build"
                        }
                    };
                    addPatchOperation("/relations/-", relation, document);
                }
                else {
                    console.log("Build " + settings.buildId + " already linked to WorkItem " + workItem.id);
                }
            }
            if (settings.updateAssignedTo === "Always" || (settings.updateAssignedTo === "Unassigned" && getWorkItemFields(workItem, (f) => f === "System.AssignedTo").length === 0)) {
                addPatchOperation("/fields/System.AssignedTo", settings.assignedTo, document);
            }
            tl.debug("Start UpdateWorkItem");
            let updatedWorkItem = yield workItemTrackingClient.updateWorkItem(null, document, workItem.id);
            console.log("WorkItem " + workItem.id + " updated");
        }
        else {
            console.log("Skipped " + workItem.fields['System.WorkItemType'] + " WorkItem: " + workItem.id);
        }
    });
}
function getWorkItemFields(workItem, predicate) {
    let result = [];
    Object.keys(workItem.fields).forEach((propertyName, index) => {
        if (predicate(propertyName)) {
            result.push(propertyName);
        }
    });
    return result;
}
function addPatchOperation(path, value, document) {
    let patchOperation = {
        from: null,
        op: VSSInterfaces_1.Operation.Add,
        path: path,
        value: value
    };
    document.push(patchOperation);
    console.log("Patch: " + patchOperation.path + " " + patchOperation.value);
}
main();

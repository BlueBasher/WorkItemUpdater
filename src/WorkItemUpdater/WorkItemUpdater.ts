import tl = require('vsts-task-lib/task');
import { Settings } from './settings';
import * as vso from 'vso-node-api';
import { IBuildApi } from 'vso-node-api/BuildApi';
import { IRequestHandler } from 'vso-node-api/interfaces/common/VsoBaseInterfaces';
import { WebApi, getHandlerFromToken } from 'vso-node-api/WebApi';
import { BuildStatus, BuildResult, BuildQueryOrder, Build } from 'vso-node-api/interfaces/BuildInterfaces';
import { IWorkItemTrackingApi } from 'vso-node-api/WorkItemTrackingApi';
import { ResourceRef, JsonPatchDocument, JsonPatchOperation, Operation } from 'vso-node-api/interfaces/common/VSSInterfaces';
import { WorkItemExpand, WorkItem, WorkItemField, WorkItemRelation } from 'vso-node-api/interfaces/WorkItemTrackingInterfaces';
import { WorkItemQueryResult } from 'vso-node-api/interfaces/WorkItemTrackingInterfaces';

async function main(): Promise<void> {
    try {
        let vstsWebApi: WebApi = getVstsWebApi();
        let settings: Settings = getSettings();

        let workItemTrackingClient: IWorkItemTrackingApi = await vstsWebApi.getWorkItemTrackingApi();

        let workItemRefs: ResourceRef[] = await getWorkItemsRefs(vstsWebApi, workItemTrackingClient, settings);
        if (!workItemRefs || workItemRefs.length === 0) {
            console.log("No workitems found to update.");
        }
        else {
            tl.debug("Loop workItemsRefs");
            await asyncForEach(workItemRefs, async (workItemRef) => {
                tl.debug("Found WorkItemRef: " + workItemRef.id);
                let workItem: WorkItem = await workItemTrackingClient.getWorkItem(parseInt(workItemRef.id), null, null, WorkItemExpand.Relations);
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

                await updateWorkItem(workItemTrackingClient, workItem, settings);
            });
            tl.debug("Finished loop workItemsRefs");
        }

        tl.setResult(tl.TaskResult.Succeeded, "");
    } catch (error) {
        tl.debug("Caught an error in main: " + error);
        tl.setResult(tl.TaskResult.Failed, error);
    }
}

async function asyncForEach(array, callback) {
    for (let index = 0; index < array.length; index++) {
        await callback(array[index], index, array)
    }
}

function getVstsWebApi() {
    let endpointUrl: string = tl.getVariable("System.TeamFoundationCollectionUri");
    let accessToken: string = tl.getEndpointAuthorizationParameter('SYSTEMVSSCONNECTION', 'AccessToken', false);
    let credentialHandler: IRequestHandler = getHandlerFromToken(accessToken);
    let webApi: WebApi = new WebApi(endpointUrl, credentialHandler);
    return webApi;
}

function getSettings(): Settings {
    let settings = new Settings();
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

async function getWorkItemsRefs(vstsWebApi: WebApi, workItemTrackingClient: IWorkItemTrackingApi, settings: Settings): Promise<ResourceRef[]> {
    if (settings.workitemsSource === 'Build') {
        console.log("Using Build as WorkItem Source");
        let buildClient: IBuildApi = await vstsWebApi.getBuildApi();
        let workItemRefs: ResourceRef[] = await buildClient.getBuildWorkItemsRefs(settings.projectId, settings.buildId);
        return workItemRefs;
    }
    else if (settings.workitemsSource === 'Query') {
        console.log("Using Query as WorkItem Source");
        var queryResult: WorkItemQueryResult = await workItemTrackingClient.queryById(
            settings.workitemsSourceQuery,
            {
                project: null,
                projectId: settings.projectId,
                team: null,
                teamId: null
            });
        var result: ResourceRef[] = [];
        queryResult.workItems.forEach((workItem) => {
            result.push({
                id: workItem.id.toString(),
                url: workItem.url
            });
        });
        return result;
    }

    return null;
}

async function updateWorkItem(workItemTrackingClient: IWorkItemTrackingApi, workItem: WorkItem, settings: Settings): Promise<void> {
    tl.debug("Updating  WorkItem: " + workItem.id);
    if (workItem.fields["System.WorkItemType"] === settings.workItemType) {
        if (settings.workItemCurrentState && settings.workItemCurrentState !== "" && settings.workItemCurrentState.split(',').indexOf(workItem.fields["System.State"]) === -1) {
            console.log("Skipped WorkItem: " + workItem.id + " State: '" + workItem.fields["System.State"] + "' => Only updating if state in '" + settings.workItemCurrentState) + "'";
            return
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
            let buildRelationUrl = "vstfs:///Build/Build/$buildId"
            let buildRelation = workItem.relations.find(r => r.url === buildRelationUrl);
            if (buildRelation === null) {
                console.log("Linking Build " + settings.buildId + " to WorkItem " + workItem.id);
                let relation: WorkItemRelation = {
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
        let updatedWorkItem = await workItemTrackingClient.updateWorkItem(null, document, workItem.id);
        console.log("WorkItem " + workItem.id + " updated");
    }
    else {
        console.log("Skipped " + workItem.fields['System.WorkItemType'] + " WorkItem: " + workItem.id);
    }
}

function getWorkItemFields(workItem, predicate): string[] {
    let result = [];
    Object.keys(workItem.fields).forEach((propertyName, index) => {
        if (predicate(propertyName)) {
            result.push(propertyName);
        }
    });
    return result;
}

function addPatchOperation(path: any, value: any, document: any[]) {
    let patchOperation: JsonPatchOperation = {
        from: null,
        op: Operation.Add,
        path: path,
        value: value
    };
    document.push(patchOperation);
    console.log("Patch: " + patchOperation.path + " " + patchOperation.value);
}

main();
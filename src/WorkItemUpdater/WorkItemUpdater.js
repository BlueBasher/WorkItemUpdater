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
const ReleaseInterfaces_1 = require("vso-node-api/interfaces/ReleaseInterfaces");
function main() {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const vstsWebApi = getVstsWebApi();
            const settings = getSettings();
            tl.debug('Get WorkItemTrackingApi');
            const workItemTrackingClient = yield vstsWebApi.getWorkItemTrackingApi();
            tl.debug('Get workItemsRefs');
            const workItemRefs = yield getWorkItemsRefs(vstsWebApi, workItemTrackingClient, settings);
            if (!workItemRefs || workItemRefs.length === 0) {
                console.log('No workitems found to update.');
            }
            else {
                tl.debug('Loop workItemsRefs');
                yield asyncForEach(workItemRefs, (workItemRef) => __awaiter(this, void 0, void 0, function* () {
                    tl.debug('Found WorkItemRef: ' + workItemRef.id);
                    const workItem = yield workItemTrackingClient.getWorkItem(parseInt(workItemRef.id), undefined, undefined, WorkItemTrackingInterfaces_1.WorkItemExpand.Relations);
                    console.log('Found WorkItem: ' + workItem.id);
                    switch (settings.updateAssignedToWith) {
                        case 'Creator': {
                            const creator = workItem.fields['System.CreatedBy'];
                            tl.debug('Using workitem creator user "' + creator + '" as assignedTo.');
                            settings.assignedTo = creator;
                            break;
                        }
                        case 'FixedUser': {
                            tl.debug('Using fixed user "' + settings.assignedTo + '" as assignedTo.');
                            break;
                        }
                        case 'Unassigned': {
                            tl.debug('Using Unassigned as assignedTo.');
                            settings.assignedTo = undefined;
                            break;
                        }
                        default: {
                            tl.debug('Setting assignedTo to requester for build "' + settings.requestedFor + '".');
                            settings.assignedTo = settings.requestedFor;
                            break;
                        }
                    }
                    yield updateWorkItem(workItemTrackingClient, workItem, settings);
                }));
                tl.debug('Finished loop workItemsRefs');
            }
            tl.setResult(tl.TaskResult.Succeeded, '');
        }
        catch (error) {
            tl.debug('Caught an error in main: ' + error);
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
    const endpointUrl = tl.getVariable('System.TeamFoundationCollectionUri');
    const accessToken = tl.getEndpointAuthorizationParameter('SYSTEMVSSCONNECTION', 'AccessToken', false);
    const credentialHandler = WebApi_1.getHandlerFromToken(accessToken);
    const webApi = new WebApi_1.WebApi(endpointUrl, credentialHandler);
    return webApi;
}
function getSettings() {
    const settings = new settings_1.Settings();
    settings.buildId = parseInt(tl.getVariable('Build.BuildId'));
    settings.projectId = tl.getVariable('System.TeamProjectId');
    settings.requestedFor = tl.getVariable('Build.RequestedFor');
    settings.workitemsSource = tl.getInput('workitemsSource');
    settings.workitemsSourceQuery = tl.getInput('workitemsSourceQuery');
    settings.allWorkItemsSinceLastRelease = tl.getBoolInput('allWorkItemsSinceLastRelease');
    settings.workItemType = tl.getInput('workItemType');
    settings.workItemState = tl.getInput('workItemState');
    settings.workItemCurrentState = tl.getInput('workItemCurrentState');
    settings.workItemKanbanLane = tl.getInput('workItemKanbanLane');
    settings.workItemKanbanState = tl.getInput('workItemKanbanState');
    settings.workItemDone = tl.getBoolInput('workItemDone');
    settings.linkBuild = tl.getBoolInput('linkBuild');
    settings.updateAssignedTo = tl.getInput('updateAssignedTo');
    settings.updateAssignedToWith = tl.getInput('updateAssignedToWith');
    settings.assignedTo = tl.getInput('assignedTo');
    settings.addTags = tl.getInput('addTags');
    if (settings.addTags) {
        settings.addTags = settings.addTags.replace(/(?:\r\n|\r|\n)/g, ';');
    }
    settings.removeTags = tl.getInput('removeTags');
    if (settings.removeTags) {
        settings.removeTags = settings.removeTags.replace(/(?:\r\n|\r|\n)/g, ';');
    }
    const releaseIdString = tl.getVariable('Release.ReleaseId');
    const definitionIdString = tl.getVariable('Release.DefinitionId');
    const definitionEnvironmentIdString = tl.getVariable('Release.DefinitionEnvironmentId');
    if (releaseIdString && releaseIdString !== ''
        && definitionIdString && definitionIdString !== ''
        && definitionEnvironmentIdString && definitionEnvironmentIdString !== '') {
        settings.releaseId = parseInt(releaseIdString);
        settings.definitionId = parseInt(definitionIdString);
        settings.definitionEnvironmentId = parseInt(definitionEnvironmentIdString);
    }
    tl.debug('BuildId ' + settings.buildId);
    tl.debug('ProjectId ' + settings.projectId);
    tl.debug('ReleaseId ' + settings.releaseId);
    tl.debug('DefinitionId ' + settings.definitionId);
    tl.debug('DefinitionEnvironmentId ' + settings.definitionEnvironmentId);
    tl.debug('requestedFor ' + settings.requestedFor);
    tl.debug('workitemsSource ' + settings.workitemsSource);
    tl.debug('workitemsSourceQuery ' + settings.workitemsSourceQuery);
    tl.debug('allWorkItemsSinceLastRelease ' + settings.allWorkItemsSinceLastRelease);
    tl.debug('workItemType ' + settings.workItemType);
    tl.debug('WorkItemState ' + settings.workItemState);
    tl.debug('workItemCurrentState ' + settings.workItemCurrentState);
    tl.debug('updateWorkItemKanbanLane ' + settings.workItemKanbanLane);
    tl.debug('WorkItemKanbanState ' + settings.workItemKanbanState);
    tl.debug('WorkItemDone ' + settings.workItemDone);
    tl.debug('updateAssignedTo ' + settings.updateAssignedTo);
    tl.debug('updateAssignedToWith ' + settings.updateAssignedToWith);
    tl.debug('assignedTo ' + settings.assignedTo);
    tl.debug('addTags ' + settings.addTags);
    tl.debug('removeTags ' + settings.removeTags);
    return settings;
}
function getWorkItemsRefs(vstsWebApi, workItemTrackingClient, settings) {
    return __awaiter(this, void 0, void 0, function* () {
        if (settings.workitemsSource === 'Build') {
            if (settings.releaseId && settings.allWorkItemsSinceLastRelease) {
                console.log('Using Release as WorkItem Source');
                const releaseClient = yield vstsWebApi.getReleaseApi();
                const deployments = yield releaseClient.getDeployments(settings.projectId, settings.definitionId, settings.definitionEnvironmentId, undefined, undefined, undefined, ReleaseInterfaces_1.DeploymentStatus.Succeeded, undefined, undefined, ReleaseInterfaces_1.ReleaseQueryOrder.Descending, 1);
                if (deployments.length > 0) {
                    const baseReleaseId = deployments[0].release.id;
                    tl.debug('Using Release ' + baseReleaseId + ' as BaseRelease for ' + settings.releaseId);
                    const releaseWorkItemRefs = yield releaseClient.getReleaseWorkItemsRefs(settings.projectId, settings.releaseId, baseReleaseId);
                    const result = [];
                    releaseWorkItemRefs.forEach((releaseWorkItem) => {
                        result.push({
                            id: releaseWorkItem.id.toString(),
                            url: releaseWorkItem.url
                        });
                    });
                    return result;
                }
            }
            console.log('Using Build as WorkItem Source');
            const buildClient = yield vstsWebApi.getBuildApi();
            const workItemRefs = yield buildClient.getBuildWorkItemsRefs(settings.projectId, settings.buildId);
            return workItemRefs;
        }
        else if (settings.workitemsSource === 'Query') {
            console.log('Using Query as WorkItem Source');
            const result = [];
            const query = yield workItemTrackingClient.getQuery(settings.projectId, settings.workitemsSourceQuery);
            if (query) {
                tl.debug('Found queryId ' + query.id + ' from QueryName ' + settings.workitemsSourceQuery);
                const queryResult = yield workItemTrackingClient.queryById(query.id, {
                    project: undefined,
                    projectId: settings.projectId,
                    team: undefined,
                    teamId: undefined
                });
                queryResult.workItems.forEach((workItem) => {
                    result.push({
                        id: workItem.id.toString(),
                        url: workItem.url
                    });
                });
            }
            else {
                tl.warning('Could not find query ' + settings.workitemsSourceQuery);
            }
            return result;
        }
        return undefined;
    });
}
function updateWorkItem(workItemTrackingClient, workItem, settings) {
    return __awaiter(this, void 0, void 0, function* () {
        tl.debug('Updating  WorkItem: ' + workItem.id);
        if (settings.workItemType.split(',').indexOf(workItem.fields['System.WorkItemType']) >= 0) {
            if (settings.workItemCurrentState && settings.workItemCurrentState !== '' && settings.workItemCurrentState.split(',').indexOf(workItem.fields['System.State']) === -1) {
                console.log('Skipped WorkItem: ' + workItem.id + ' State: "' + workItem.fields['System.State'] + '" => Only updating if state in "' + settings.workItemCurrentState) + '"';
                return;
            }
            console.log('Updating WorkItem ' + workItem.id);
            const document = [];
            const kanbanLane = getWorkItemFields(workItem, (f) => f.endsWith('Kanban.Lane'));
            tl.debug('Found KanbanLane: ' + kanbanLane);
            const kanbanColumn = getWorkItemFields(workItem, (f) => f.endsWith('Kanban.Column'));
            tl.debug('Found KanbanColumn: ' + kanbanColumn);
            const kanbanColumnDone = getWorkItemFields(workItem, (f) => f.endsWith('Kanban.Column.Done'));
            tl.debug('Found KanbanColumnDone: ' + kanbanColumnDone);
            if (settings.workItemState && settings.workItemState !== '') {
                addPatchOperation('/fields/System.State', settings.workItemState, document);
            }
            if (settings.workItemKanbanLane && settings.workItemKanbanLane !== '' && kanbanLane.length > 0) {
                kanbanLane.forEach((lane, index) => {
                    addPatchOperation('/fields/' + lane, settings.workItemKanbanLane, document);
                });
            }
            if (settings.workItemKanbanState && settings.workItemKanbanState !== '' && kanbanColumn.length > 0) {
                kanbanColumn.forEach((column, index) => {
                    addPatchOperation('/fields/' + column, settings.workItemKanbanState, document);
                });
            }
            if (kanbanColumnDone.length > 0) {
                kanbanColumnDone.forEach((columnDone, index) => {
                    addPatchOperation('/fields/' + columnDone, settings.workItemDone, document);
                });
            }
            if (settings.linkBuild) {
                const buildRelationUrl = 'vstfs:///Build/Build/$buildId';
                const buildRelation = workItem.relations.find(r => r.url === buildRelationUrl);
                if (buildRelation === null) {
                    console.log('Linking Build ' + settings.buildId + ' to WorkItem ' + workItem.id);
                    const relation = {
                        rel: 'ArtifactLink',
                        url: buildRelationUrl,
                        attributes: {
                            name: 'Build'
                        }
                    };
                    addPatchOperation('/relations/-', relation, document);
                }
                else {
                    console.log('Build ' + settings.buildId + ' already linked to WorkItem ' + workItem.id);
                }
            }
            if (settings.updateAssignedTo === 'Always' || (settings.updateAssignedTo === 'Unassigned' && getWorkItemFields(workItem, (f) => f === 'System.AssignedTo').length === 0)) {
                let operation = VSSInterfaces_1.Operation.Add;
                if (!settings.assignedTo) {
                    operation = VSSInterfaces_1.Operation.Remove;
                }
                addPatchOperation('/fields/System.AssignedTo', settings.assignedTo, document, operation);
            }
            if (settings.addTags || settings.removeTags) {
                const newTags = [];
                const removeTags = settings.removeTags ? settings.removeTags.split(';') : [];
                if (workItem.fields['System.Tags']) {
                    tl.debug('Existing tags: ' + workItem.fields['System.Tags']);
                    workItem.fields['System.Tags'].split(';').forEach((tag) => {
                        if (removeTags.find(e => e.toLowerCase() === tag.trim().toLowerCase())) {
                            tl.debug('Removing tag: ' + tag);
                        }
                        else {
                            newTags.push(tag.trim());
                        }
                    });
                }
                const addTags = settings.addTags ? settings.addTags.split(';') : [];
                addTags.forEach((tag) => {
                    if (!newTags.find(e => e.toLowerCase() === tag.toLowerCase())) {
                        tl.debug('Adding tag: ' + tag);
                        newTags.push(tag.trim());
                    }
                });
                addPatchOperation('/fields/System.Tags', newTags.join('; '), document);
            }
            tl.debug('Start UpdateWorkItem');
            const updatedWorkItem = yield workItemTrackingClient.updateWorkItem(undefined, document, workItem.id);
            console.log('WorkItem ' + workItem.id + ' updated');
        }
        else {
            console.log('Skipped ' + workItem.fields['System.WorkItemType'] + ' WorkItem: ' + workItem.id);
        }
    });
}
function getWorkItemFields(workItem, predicate) {
    const result = [];
    Object.keys(workItem.fields).forEach((propertyName, index) => {
        if (predicate(propertyName)) {
            result.push(propertyName);
        }
    });
    return result;
}
function addPatchOperation(path, value, document, operation = VSSInterfaces_1.Operation.Add) {
    const patchOperation = {
        from: undefined,
        op: operation,
        path: path,
        value: value
    };
    document.push(patchOperation);
    console.log('Patch: ' + patchOperation.path + ' ' + patchOperation.value);
}
main();

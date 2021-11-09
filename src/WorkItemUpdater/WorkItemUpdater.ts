import tl = require('azure-pipelines-task-lib/task');
import { Settings } from './settings';
import * as moment from 'moment'
import * as vso from 'azure-devops-node-api';
import { IBuildApi } from 'azure-devops-node-api/BuildApi';
import { IRequestHandler } from 'azure-devops-node-api/interfaces/common/VsoBaseInterfaces';
import { WebApi, getHandlerFromToken } from 'azure-devops-node-api/WebApi';
import { BuildStatus, BuildResult, BuildQueryOrder, Build } from 'azure-devops-node-api/interfaces/BuildInterfaces';
import { IWorkItemTrackingApi } from 'azure-devops-node-api/WorkItemTrackingApi';
import { ResourceRef, JsonPatchDocument, JsonPatchOperation, Operation } from 'azure-devops-node-api/interfaces/common/VSSInterfaces';
import { WorkItemExpand, WorkItem, WorkItemField, WorkItemRelation, QueryHierarchyItem, FieldType } from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';
import { WorkItemQueryResult } from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';
import { IReleaseApi } from 'azure-devops-node-api/ReleaseApi';
import { DeploymentStatus, ReleaseQueryOrder } from 'azure-devops-node-api/interfaces/ReleaseInterfaces';

async function main(): Promise<void> {
    try {
        const vstsWebApi: WebApi = getVstsWebApi();
        const settings: Settings = getSettings();

        tl.debug('Get WorkItemTrackingApi');
        const workItemTrackingClient: IWorkItemTrackingApi = await vstsWebApi.getWorkItemTrackingApi();

        tl.debug('Get workItemsRefs');
        const workItemRefs: ResourceRef[] = await getWorkItemsRefs(vstsWebApi, workItemTrackingClient, settings);
        let numberOfUpdateWorkitems = 0;
        if (!workItemRefs || workItemRefs.length === 0) {
            console.log('No workitems found to update.');
        }
        else {
            tl.debug('Loop workItemsRefs');
            await asyncForEach(workItemRefs, async (workItemRef: ResourceRef) => {
                tl.debug('Found WorkItemRef: ' + workItemRef.id);
                const workItem: WorkItem = await workItemTrackingClient.getWorkItem(parseInt(workItemRef.id), undefined, undefined, WorkItemExpand.Relations);
                console.log('Found WorkItem: ' + workItem.id);

                switch (settings.updateAssignedToWith) {
                    case 'Creator': {
                        settings.assignedTo = workItem.fields['System.CreatedBy'];
                        tl.debug('Using workitem creator user "' + settings.assignedTo + '" as assignedTo.');
                        break;
                    }
                    case 'ActivatedBy': {
                        settings.assignedTo = workItem.fields['Microsoft.VSTS.Common.ActivatedBy'];
                        tl.debug('Using workitem activator user "' + settings.assignedTo + '" as assignedTo.');
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

                if (await updateWorkItem(workItemTrackingClient, workItem, settings)) {
                    numberOfUpdateWorkitems++;
                }
            });
            tl.debug('Finished loop workItemsRefs');
        }

        if (numberOfUpdateWorkitems == 0
            && settings.failTaskIfNoWorkItemsAvailable) {
            tl.setResult(tl.TaskResult.Failed, 'Found no workitems to update.');
        } else {
            tl.setResult(tl.TaskResult.Succeeded, '');
        }
    } catch (error) {
        tl.debug('Caught an error in main: ' + error);
        tl.setResult(tl.TaskResult.Failed, error);
    }
}

async function asyncForEach(array: any[], callback: any) {
    for (let index = 0; index < array.length; index++) {
        await callback(array[index], index, array);
    }
}

function getVstsWebApi() {
    const endpointUrl: string = tl.getVariable('System.TeamFoundationCollectionUri');
    const accessToken: string = tl.getEndpointAuthorizationParameter('SYSTEMVSSCONNECTION', 'AccessToken', false);
    const credentialHandler: IRequestHandler = getHandlerFromToken(accessToken);
    const webApi: WebApi = new WebApi(endpointUrl, credentialHandler);
    return webApi;
}

function getSettings(): Settings {
    const settings = new Settings();
    settings.buildId = parseInt(tl.getVariable('Build.BuildId'));
    settings.projectId = tl.getVariable('System.TeamProjectId');
    settings.requestedFor = tl.getVariable('Build.RequestedFor');
    settings.workitemsSource = tl.getInput('workitemsSource');
    settings.workitemsSourceQuery = tl.getInput('workitemsSourceQuery');
    settings.allWorkItemsSinceLastRelease = tl.getBoolInput('allWorkItemsSinceLastRelease');
    settings.workItemType = tl.getInput('workItemType');
    settings.workitemLimit = parseInt(tl.getInput('workitemLimit'));
    settings.workItemState = tl.getInput('workItemState');
    settings.workItemCurrentState = tl.getInput('workItemCurrentState');
    settings.workItemKanbanLane = tl.getInput('workItemKanbanLane');
    settings.workItemKanbanState = tl.getInput('workItemKanbanState');
    settings.workItemDone = tl.getBoolInput('workItemDone');
    settings.linkBuild = tl.getBoolInput('linkBuild');
    settings.updateAssignedTo = tl.getInput('updateAssignedTo');
    settings.updateAssignedToWith = tl.getInput('updateAssignedToWith');
    settings.assignedTo = tl.getInput('assignedTo');
    settings.comment = tl.getInput('comment');
    settings.updateFields = tl.getInput('updateFields');
    settings.bypassRules = tl.getBoolInput('bypassRules');
    settings.failTaskIfNoWorkItemsAvailable = tl.getBoolInput('failTaskIfNoWorkItemsAvailable');

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
    tl.debug('workitemLimit ' + settings.workitemLimit);
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
    tl.debug('updateFields ' + settings.updateFields);
    tl.debug('comment ' + settings.comment);
    tl.debug('removeTags ' + settings.removeTags);
    tl.debug('bypassRules ' + settings.bypassRules);
    tl.debug('failTaskIfNoWorkItemsAvailable ' + settings.failTaskIfNoWorkItemsAvailable);

    return settings;
}

async function getWorkItemsRefs(vstsWebApi: WebApi, workItemTrackingClient: IWorkItemTrackingApi, settings: Settings): Promise<ResourceRef[]> {
    if (settings.workitemsSource === 'Build') {
        if (settings.releaseId && settings.allWorkItemsSinceLastRelease) {
            console.log('Using Release as WorkItem Source');
            const releaseClient: IReleaseApi = await vstsWebApi.getReleaseApi();
            const deployments = await releaseClient.getDeployments(settings.projectId, settings.definitionId, settings.definitionEnvironmentId, undefined, undefined, undefined, DeploymentStatus.Succeeded, undefined, undefined, ReleaseQueryOrder.Descending, 1);
            if (deployments.length > 0) {
                const baseReleaseId = deployments[0].release.id;
                tl.debug('Using Release ' + baseReleaseId + ' as BaseRelease for ' + settings.releaseId);
                const releaseWorkItemRefs = await releaseClient.getReleaseWorkItemsRefs(settings.projectId, settings.releaseId, baseReleaseId);
                const result: ResourceRef[] = [];
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
        const buildClient: IBuildApi = await vstsWebApi.getBuildApi();
        const workItemRefs: ResourceRef[] = await buildClient.getBuildWorkItemsRefs(settings.projectId, settings.buildId, settings.workitemLimit);
        return workItemRefs;
    }
    else if (settings.workitemsSource === 'Query') {
        console.log('Using Query as WorkItem Source');
        const result: ResourceRef[] = [];
        const query: QueryHierarchyItem = await workItemTrackingClient.getQuery(settings.projectId, settings.workitemsSourceQuery);
        if (query) {
            tl.debug('Found queryId ' + query.id + ' from QueryName ' + settings.workitemsSourceQuery);
            const queryResult: WorkItemQueryResult = await workItemTrackingClient.queryById(
                query.id,
                {
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
}

async function updateWorkItem(workItemTrackingClient: IWorkItemTrackingApi, workItem: WorkItem, settings: Settings): Promise<boolean> {
    tl.debug('Updating  WorkItem: ' + workItem.id);
    if (settings.workItemType.split(',').indexOf(workItem.fields['System.WorkItemType']) >= 0) {
        if (settings.workItemCurrentState && settings.workItemCurrentState !== '' && settings.workItemCurrentState.split(',').indexOf(workItem.fields['System.State']) === -1) {
            console.log('Skipped WorkItem: ' + workItem.id + ' State: "' + workItem.fields['System.State'] + '" => Only updating if state in "' + settings.workItemCurrentState) + '"';
            return false;
        }

        console.log('Updating WorkItem ' + workItem.id);

        const document: any[] = [];

        const kanbanLane = getWorkItemFields(workItem, (f: string) => f.endsWith('Kanban.Lane'));
        tl.debug('Found KanbanLane: ' + kanbanLane);
        const kanbanColumn = getWorkItemFields(workItem, (f: string) => f.endsWith('Kanban.Column'));
        tl.debug('Found KanbanColumn: ' + kanbanColumn);
        const kanbanColumnDone = getWorkItemFields(workItem, (f: string) => f.endsWith('Kanban.Column.Done'));
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

        if (settings.workItemDone && kanbanColumnDone.length > 0) {
            kanbanColumnDone.forEach((columnDone, index) => {
                addPatchOperation('/fields/' + columnDone, settings.workItemDone, document);
            });
        }

        if (settings.linkBuild) {
            const buildRelationUrl = `vstfs:///Build/Build/${settings.buildId}`;
            const buildRelation = !workItem.relations || workItem.relations.find(r => r.url === buildRelationUrl);
            if (buildRelation === undefined) {
                console.log('Linking Build ' + settings.buildId + ' to WorkItem ' + workItem.id);
                const relation: WorkItemRelation = {
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

        if (settings.updateAssignedTo === 'Always' || (settings.updateAssignedTo === 'Unassigned' && getWorkItemFields(workItem, (f: string) => f === 'System.AssignedTo').length === 0)) {
            let operation: Operation = Operation.Add;
            if (!settings.assignedTo) {
                operation = Operation.Remove;
            }

            addPatchOperation('/fields/System.AssignedTo', settings.assignedTo, document, operation);
        }

        if (settings.addTags || settings.removeTags) {
            let operation: Operation = Operation.Add;
            const newTags: string[] = [];

            const removeTags: string[] = settings.removeTags ? settings.removeTags.split(';') : [];
            
            if (removeTags.length > 0) operation = Operation.Replace;
            
            if (workItem.fields['System.Tags']) {
                tl.debug('Existing tags: ' + workItem.fields['System.Tags']);
                workItem.fields['System.Tags'].split(';').forEach((tag: string) => {
                    if (removeTags.find(e => e.toLowerCase() === tag.trim().toLowerCase())) {
                        tl.debug('Removing tag: ' + tag);
                    }
                    else {
                        newTags.push(tag.trim());
                    }
                });
            }

            const addTags: string[] = settings.addTags ? settings.addTags.split(';') : [];
            addTags.forEach((tag) => {
                if (!newTags.find(e => e.toLowerCase() === tag.toLowerCase())) {
                    tl.debug('Adding tag: ' + tag);
                    newTags.push(tag.trim());
                }
            });

            addPatchOperation('/fields/System.Tags', newTags.join('; '), document, operation);
        }

        if (settings.updateFields) {
            const updateFields: string[] = settings.updateFields.split(/\r?\n/);
            updateFields.forEach((updateField) => {
                const commaIndex = updateField.indexOf(',');
                if (commaIndex >= 0) {
                    const fieldName = updateField.substring(0, commaIndex);
                    let fieldValue = updateField.substring(commaIndex + 1);
                    const workItemField = workItem.fields[fieldName] as WorkItemField;
                    if (workItemField) {
                        if (workItemField.type === FieldType.DateTime) {
                            // wrap date values with moment in a hope to correct an invalid date
                            const date = moment(fieldValue);
                            if (date.isValid()) {
                                fieldValue = date.format()
                            } else {
                                console.log('Skipped updating' + fieldName + ' to ' + fieldValue + ' is not a valid date');
                            }
                        }
                    }
                    console.log('Updating' + fieldName + ' to ' + fieldValue);
                    addPatchOperation('/fields/' + fieldName, fieldValue, document);
                }
            });
        }

        tl.debug('Start UpdateWorkItem');


        if (document.length === 0 && !settings.comment) {
            // workItemTrackingClient.updateWorkItem fails if there is not patch operation
            console.log('No update for WorkItem ' + workItem.id);
            return false;
        }

        if (document.length > 0) {
            const updatedWorkItem = await workItemTrackingClient.updateWorkItem(
                undefined,
                document,
                workItem.id,
                undefined,
				false,
                settings.bypassRules);
            console.log('WorkItem ' + workItem.id + ' updated');
        }
        if (settings.comment) {
            const addCommentToWorkItem = await workItemTrackingClient.addComment(
                { text: settings.comment },
                settings.projectId,
                workItem.id
            );
            console.log('WorkItem ' + workItem.id + ' comment added: "' + settings.comment + '"');
        }
        return true;
    }
    else {
        console.log('Skipped ' + workItem.fields['System.WorkItemType'] + ' WorkItem: ' + workItem.id);
    }

    return false;
}

function getWorkItemFields(workItem: any, predicate: any): string[] {
    const result: string[] = [];
    Object.keys(workItem.fields).forEach((propertyName, index) => {
        if (predicate(propertyName)) {
            result.push(propertyName);
        }
    });
    return result;
}

function addPatchOperation(path: any, value: any, document: any[], operation: Operation = Operation.Add) {
    const patchOperation: JsonPatchOperation = {
        from: undefined,
        op: operation,
        path: path,
        value: value
    };
    document.push(patchOperation);
    console.log('Patch: ' + patchOperation.path + ' ' + patchOperation.value);
}

main();

export class Settings {
    public buildId: number;
    public projectId: string;
    public releaseId: number | null;
    public definitionId: number | null;
    public definitionEnvironmentId: number | null;
    public requestedFor: string;
    public workitemsSource: string;
    public workitemsSourceQuery: string;
    public workItemType: string;
    public allWorkItemsSinceLastRelease: boolean;
    public workItemState: string;
    public workItemCurrentState: string;
    public workItemKanbanLane: string;
    public workItemKanbanState: string;
    public workItemDone: boolean;
    public linkBuild: boolean;
    public updateAssignedTo: string;
    public updateAssignedToWith: string;
    public assignedTo: string;
    public addTags: string;
    public removeTags: string;
    public updateFields: string;
    public bypassRules: boolean;
    public failTaskIfNoWorkItemsAvailable: boolean;
}
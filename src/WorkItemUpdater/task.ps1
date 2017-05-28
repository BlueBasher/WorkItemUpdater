#
# WorkItemUpdater.ps1
#
[CmdletBinding(DefaultParameterSetName = 'None')]
param()

$directory = [System.IO.Path]::GetFullPath("$PSScriptRoot\")
$newtonsoftDll = [System.IO.Path]::Combine($directory, "Newtonsoft.Json.dll")
$httpFormatingDll = [System.IO.Path]::Combine($directory, "System.Net.Http.Formatting.dll")
$onAssemblyResolve = [System.ResolveEventHandler]{
    param($sender, $e)

    if ($e.Name -like 'Newtonsoft.Json, *') {
        Write-Verbose "Resolving '$($e.Name)'"
        return [System.Reflection.Assembly]::LoadFrom($newtonsoftDll)
    }
	else
	{
		if ($e.Name -like 'System.Net.Http.Formatting, *') {
			Write-Verbose "Resolving '$($e.Name)'"
			return [System.Reflection.Assembly]::LoadFrom($httpFormatingDll)
		}
	}
    return $null
}
[System.AppDomain]::CurrentDomain.add_AssemblyResolve($onAssemblyResolve)

Import-Module .\ps_modules\VstsTaskSdk\VstsTaskSdk.psm1 -Verbose:$true

Trace-VstsEnteringInvocation $MyInvocation

function Update-WorkItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.TeamFoundation.WorkItemTracking.WebApi.WorkItemTrackingHttpClient]$workItemTrackingHttpClient,
        [Parameter(Mandatory = $true)]
        [int]$workItemId,
        [Parameter(Mandatory = $true)]
        [int]$buildId,
        [Parameter(Mandatory = $true)]
        [string]$workItemType,
        [Parameter(Mandatory = $true)]
        [string]$workItemState,
        [Parameter(Mandatory = $false)]
        [string]$workItemKanbanState,
        [Parameter(Mandatory = $true)]
        [string]$workItemDone,
        [Parameter(Mandatory = $true)]
        [bool]$linkBuild,
        [Parameter(Mandatory = $true)]
        [string]$requestedFor,
        [Parameter(Mandatory = $true)]
        [string]$updateAssignedTo)

	Write-VstsTaskDebug -Message "Found WorkItemRef: $($workItemId)"
	$workItem = $workItemTrackingHttpClient.GetWorkItemAsync($workItemId, $null, $null, [Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models.WorkItemExpand]::Relations).Result
	Write-VstsTaskDebug -Message "Found WorkItem: $($workItem.Id)"
	if ($workItem.Fields["System.WorkItemType"] -eq $workItemType)
	{
		Write-Host "Updating WorkItem $($workItem.Id)"

		$kanbanColumn = $workItem.Fields.Keys | Where-Object { $_.EndsWith("Kanban.Column") }
		Write-VstsTaskDebug -Message "Found KanbanColumn: $($kanbanColumn)"
		$kanbanField = "/fields/$($kanbanColumn)"
		$kanbanDoneField = "/fields/$($kanbanColumn).Done"

		$patch = New-Object Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchDocument
		$columnOperation = New-Object Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchOperation
		$columnOperation.Operation = [Microsoft.VisualStudio.Services.WebApi.Patch.Operation]::Add
		$columnOperation.Path = "/fields/System.State"
		$columnOperation.Value = $workItemState
		$patch.Add($columnOperation)
		Write-VstsTaskDebug -Message "Patch: $($columnOperation.Path) $($columnOperation.Value)"

		if ($workItemKanbanState -ne "")
		{
			$columnDoneOperation = New-Object Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchOperation
			$columnDoneOperation.Operation = [Microsoft.VisualStudio.Services.WebApi.Patch.Operation]::Add
			$columnDoneOperation.Path = $kanbanField
			$columnDoneOperation.Value = $workItemKanbanState
			$patch.Add($columnDoneOperation)
			Write-VstsTaskDebug -Message "Patch: $($columnDoneOperation.Path) $($columnDoneOperation.Value)"
		}

		$columnDoneOperation = New-Object Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchOperation
		$columnDoneOperation.Operation = [Microsoft.VisualStudio.Services.WebApi.Patch.Operation]::Add
		$columnDoneOperation.Path = $kanbanDoneField
		$columnDoneOperation.Value = $workItemDone
		$patch.Add($columnDoneOperation)
		Write-VstsTaskDebug -Message "Patch: $($columnDoneOperation.Path) $($columnDoneOperation.Value)"

		if ($linkBuild -eq $true)
		{
			$buildRelationUrl = "vstfs:///Build/Build/$buildId"
			$buildRelation = $workItem.Relations | Where-Object { $_.Url -eq $buildRelationUrl }
			if ($buildRelation -eq $null) {
				Write-Host "Linking Build $($buildId) to WorkItem $($workItem.Id)"
				$linkBuildOperation = New-Object Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchOperation
				$linkBuildOperation.Operation = [Microsoft.VisualStudio.Services.WebApi.Patch.Operation]::Add
				$linkBuildOperation.Path = "/relations/-"
				$linkBuildOperation.Value = @{
						Rel = "ArtifactLink"
						Url = $buildRelationUrl
						Attributes = @{
							name = "Build"
						}
					}
				$patch.Add($linkBuildOperation)
				Write-VstsTaskDebug -Message "Patch: $($linkBuildOperation.Path) $($buildRelationUrl)"
			}
			else {
				Write-Host "Build $($buildId) already linked to WorkItem $($workItem.Id)"
			}
		}
		if ($updateAssignedTo -eq "Always" -or ($updateAssignedTo -eq "Unassigned" -and $workItem.Fields.ContainsKey("System.AssignedTo") -eq $false))
		{
			$assignedToOperation = New-Object Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchOperation
			$assignedToOperation.Operation = [Microsoft.VisualStudio.Services.WebApi.Patch.Operation]::Add
			$assignedToOperation.Path = "/fields/System.AssignedTo"
			$assignedToOperation.Value = $requestedFor
			$patch.Add($assignedToOperation)
			Write-VstsTaskDebug -Message "Patch: $($assignedToOperation.Path) $($assignedToOperation.Value)"
		}

		$updatedWorkItem = $workItemTrackingHttpClient.UpdateWorkItemAsync($patch, $workItem.Id).Result

		Write-Host "WorkItem $($updatedWorkItem.Id) updated to $($workItemState) $($workItemKanbanState) $($workItemDone)"
	}
	else
	{
		Write-VstsTaskDebug -Message "Skipped $($workItem.Fields['System.WorkItemType']) WorkItem: $($workItem.Id)"
	}
}

try {
    $buildId = Get-VstsTaskVariable -Name "Build.BuildId"
	$projectId = Get-VstsTaskVariable -Name "System.TeamProjectId"
	$requestedFor = Get-VstsTaskVariable -Name "Build.RequestedFor"
	$workItemType = Get-VstsInput -Name "workItemType"
	$workItemState = Get-VstsInput -Name "workItemState"
	$workItemKanbanState = Get-VstsInput -Name "workItemKanbanState"
	$workItemDone = Get-VstsInput -Name "workItemDone" -AsBool 
	$linkBuild = Get-VstsInput -Name "linkBuild" -AsBool
	$updateAssignedTo = Get-VstsInput -Name "updateAssignedTo"

    Write-VstsTaskDebug -Message "BuildId $buildId"
    Write-VstsTaskDebug -Message "ProjectId $projectId"
    Write-VstsTaskDebug -Message "requestedFor $requestedFor"
    Write-VstsTaskDebug -Message "workItemType $workItemType"
    Write-VstsTaskDebug -Message "WorkItemState $workItemState"
    Write-VstsTaskDebug -Message "WorkItemKanbanState $workItemKanbanState"
    Write-VstsTaskDebug -Message "WorkItemDone $workItemDone"
    Write-VstsTaskDebug -Message "updateAssignedTo $updateAssignedTo"

	$workItemTrackingHttpClient = Get-VssHttpClient -TypeName Microsoft.TeamFoundation.WorkItemTracking.WebApi.WorkItemTrackingHttpClient
    $buildHttpClient = Get-VssHttpClient -TypeName Microsoft.TeamFoundation.Build.WebApi.BuildHttpClient
	$workItemsRefs = $buildHttpClient.GetBuildWorkItemsRefsAsync($projectId, $buildId).Result
	foreach ($workItemRef in $workItemsRefs)
	{
		Update-WorkItem -workItemTrackingHttpClient $workItemTrackingHttpClient `
        	-workItemId $workItemRef.Id `
        	-buildId $buildId `
			-workItemType $workItemType `
			-workItemState $workItemState `
			-workItemKanbanState $workItemKanbanState `
			-workItemDone $workItemDone `
			-linkBuild $linkBuild `
			-requestedFor $requestedFor `
			-updateAssignedTo $updateAssignedTo
	}
}
catch {
 	Write-Host $_.Exception.Message
 	Write-Host $_.Exception.StackTrace
	Write-VstsSetResult -Result "Error updating workitems"
}
finally {
    Trace-VstsLeavingInvocation $MyInvocation
}
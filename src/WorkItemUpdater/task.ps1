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

    Write-Verbose "Unable to resolve assembly name '$($e.Name)'"
    return $null
}
[System.AppDomain]::CurrentDomain.add_AssemblyResolve($onAssemblyResolve)

Import-Module .\ps_modules\VstsTaskSdk\VstsTaskSdk.psm1 -Verbose:$true

Trace-VstsEnteringInvocation $MyInvocation

try {
    $buildId = Get-VstsTaskVariable -Name "Build.BuildId"
	$projectId = Get-VstsTaskVariable -Name "System.TeamProjectId"
	$workItemType = Get-VstsInput -Name "workItemType"
	$workItemState = Get-VstsInput -Name "workItemState"
	$workItemDone = Get-VstsInput -Name "workItemDone" -AsBool 

    Write-VstsTaskDebug -Message "BuildId $buildId"
    Write-VstsTaskDebug -Message "ProjectId $projectId"
    Write-VstsTaskDebug -Message "workItemType $workItemType"
    Write-VstsTaskDebug -Message "WorkItemState $workItemState"
    Write-VstsTaskDebug -Message "WorkItemDone $workItemDone"

	$workItemTrackingHttpClient = Get-VssHttpClient -TypeName Microsoft.TeamFoundation.WorkItemTracking.WebApi.WorkItemTrackingHttpClient
    $buildHttpClient = Get-VssHttpClient -TypeName Microsoft.TeamFoundation.Build.WebApi.BuildHttpClient
	$workItemsRefs = $buildHttpClient.GetBuildWorkItemsRefsAsync($projectId, $buildId).Result
	foreach ($workItemRef in $workItemsRefs)
	{
		Write-VstsTaskDebug -Message "Found WorkItemRef: $($workItemRef.Id)"
		$workItem = $workItemTrackingHttpClient.GetWorkItemAsync($workItemRef.Id).Result
		Write-VstsTaskDebug -Message "Found WorkItem: $($workItem.Id)"
		if ($workItem.Fields["System.WorkItemType"] -eq $workItemType)
		{
			$kanbanColumn = $workItem.Fields.Keys | where { $_.EndsWith("Kanban.Column") }
			Write-VstsTaskDebug -Message "Found KanbanColumn: $($kanbanColumn)"
			$kanbanField = "/fields/$($kanbanColumn).Done"

			$patch = New-Object Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchDocument
			$columnOperation = New-Object Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchOperation
			$columnOperation.Operation = 0
			$columnOperation.Path = "/fields/System.State"
			$columnOperation.Value = $workItemState
			$patch.Add($columnOperation)
			Write-VstsTaskDebug -Message "Patch: $($columnOperation.Path) $($columnOperation.Value)"

			$columnDoneOperation = New-Object Microsoft.VisualStudio.Services.WebApi.Patch.Json.JsonPatchOperation
			$columnDoneOperation.Operation = 0
			$columnDoneOperation.Path = $kanbanField
			$columnDoneOperation.Value = $workItemDone
			$patch.Add($columnDoneOperation)
			Write-VstsTaskDebug -Message "Patch: $($columnDoneOperation.Path) $($columnDoneOperation.Value)"

			Write-Host "Updating WorkItem $($workItem.Id)"
			$updateResult = $workItemTrackingHttpClient.UpdateWorkItemAsync($patch, $workItem.Id).Result
			Write-Host "WorkItem $($workItem.Id) updated to $($workItemState) $($workItemDone)"
		}
		else
		{
			Write-VstsTaskDebug -Message "Skipped $($workItem.Fields['System.WorkItemType']) WorkItem: $($workItem.Id)"
		}
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
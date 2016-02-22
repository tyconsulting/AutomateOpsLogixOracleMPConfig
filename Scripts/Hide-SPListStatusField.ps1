$IsSharePointOnlineSite = $true
$SiteUrl = "https://tyconsultingnet.sharepoint.com/"
$ListTitle = "OpsLogix Oracle Perf Rule Request - Demo"
$credential = Get-Credential -Message "Please enter the user name and password for the sharepoint credential"

Write-Output "Hiding the 'Automation Status' field from the new and edit form"
$UpdateFieldVisibility = Set-SPListFieldVisibility -SiteUrl $SiteUrl -Credential $credential -IsSharePointOnlineSite $IsSharePointOnlineSite -ListTitle $ListTitle -FieldName "Automation Status" -ShowInEditForm $false -ShowInNewForm $false -ShowInDisplayForm $true -Verbose

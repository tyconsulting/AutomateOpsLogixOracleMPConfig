param ([object]$WebHookData)

#region defining variables
$OpsMgrSConnectionName = "OM_OpsLogixDemo"
$SharePointSDKConnectionName = 'SP_OpsLogixDemo'
$MPName = "TYANG.OpsLogix.Oracle.Demo"
$MPDisplayName = "TYANG OpsLogix Oracle Demo"
$LocaleId="ENU"
$IncreaseMPVersion = $true
#endregion

#region process inputs from webhook data
Write-Verbose "Processing inputs from webhook data."
$WebhookName    =   $WebhookData.WebhookName
Write-Verbose "Webhook name: '$WebhookName'"
$WebhookHeaders =   $WebhookData.RequestHeader
$WebhookBody    =   $WebhookData.RequestBody
Write-Verbose "Webhook body:"
Write-Verbose $WebhookBody
$Inputs = ConvertFrom-JSON $webhookdata.RequestBody
$DisplayName = $Inputs.DisplayName
$Description = $Inputs.Description
$Enabled = $Inputs.Enabled
$CounterName = $Inputs.CounterName
$Query = $Inputs.Query
$ReturnColumnName = $Inputs.ReturnColumnName
$Target = $Inputs.Target
$QueryInterval = $Inputs.QueryInterval
$ListItemId = $Inputs.ListItemId
$ListTitle = $Inputs.ListTitle
Write-Verbose "DisplayName: '$DisplayName'"
Write-Verbose "Description: '$Description'"
Write-Verbose "Enabled: '$Enabled'"
Write-Verbose "CounterName: '$CounterName'"
Write-Verbose "Query: '$Query'"
Write-Verbose "ReturnColumnName: '$ReturnColumnName'"
Write-Verbose "Target: '$Target'"
Write-Verbose "QueryInterval: '$QueryInterval'"
Write-Verbose "ListItemId: '$ListItemId'"
Write-Verbose "ListTitle: '$ListTitle'"
#endregion

#Get OpsMgr connection
$OpsMgrSConnectionName = $OpsMgrSConnectionName
Write-Verbose "Getting OpsMgr SDK connection object."
$OpsMgrSDKConn = Get-AutomationConnection -Name $OpsMgrSConnectionName
Write-Verbose "Management Server: '$($OpsMgrSDKConn.ComputerName)'."

#Get SharePoint Connection
If ($ListItemID -and $ListTitle)
{
    $SPConnection = Get-AutomationConnection $SharePointSDKConnectionName
    Write-Verbose "SharePoint Site URL: '$($SPConnection.SharePointSiteURL)'"
    Write-Verbose "SharePoint List Name: '$ListTitle'"
}

#Connect to MG
Write-Verbose "Connecting to Management Group via SDK $($SDKConnection.ComputerName)`..."
$MG = Connect-OMManagementGroup -SDKConnection $OpsMgrSDKConn

#Get the unsealed MP
Write-Verbose "Getting destination MP '$MPName'..."
$strMPquery = "Name = '$MPName'"
$mpCriteria = New-Object  Microsoft.EnterpriseManagement.Configuration.ManagementPackCriteria($strMPquery)
$MP = $MG.GetManagementPacks($mpCriteria)[0]

If ($MP)
{
    #MP found, now check if it is sealed
    Write-Verbose "Found destination MP '$MPName', the MP display name is '$($MP.DisplayName)'. MP Sealed: $($MP.Sealed)"
    Write-Verbose $MP.GetType()
    If ($MP.sealed)
    {
        Write-Error 'Unable to save to the management pack specified. It is sealed. Please specify an unsealed MP.'
        return
    }
} else {
    Write-Verbose 'The management pack specified cannot be found. please make sure the correct name is specified.'
    $CreateMP = New-OMManagementPack -SDKConnection $OpsMgrSDKConn -Name $MPName -DisplayName $MPDisplayName -Version "1.0.0.0"
    If ($CreateMP -eq $true)
    {
        Write-Verbose "Unsealed MP 'MPName' Successfully Created."
        $MP = $MG.GetManagementPacks($mpCriteria)[0]
    } else {
        Write-Error "Unable to create MP 'MPName'. Unable to continue. exiting..."
        Return
    }
}

#Get the target class
Switch ($Target.ToLower())
{
    'instance' {$TargetClass = 'OpsLogix.IMP.Oracle.Instance'}
    'control file' {$TargetClass = 'OpsLogix.IMP.Oracle.ControlFile'}
    'data file' {$TargetClass = 'OpsLogix.IMP.Oracle.DataFile'}
    'redo log group' {$TargetClass = 'OpsLogix.IMP.Oracle.RedoLogGroup'}
    'redo log file' {$TargetClass = 'OpsLogix.IMP.Oracle.RedoLogFile'}
    'table space' {$TargetClass = 'OpsLogix.IMP.Oracle.TableSpace'}
}


#Get the template
Write-Verbose "Getting the OpsLogix Oracle Performance Collection Rule monitoring template..."
$strTemplatequery = "Name = 'OpsLogix.IMP.Oracle.Config.2012.SQL.Query.Template'"
$TemplateCriteria = New-object Microsoft.EnterpriseManagement.Configuration.ManagementPackTemplateCriteria($strTemplatequery)
$OpsLogixOraclePerfTemplate = $MG.GetMonitoringTemplates($TemplateCriteria)[0]
if (!$OpsLogixOraclePerfTemplate)
{
    Write-Error "The Opslogix Oracle Performance Collection Rule Monitoring Template cannot be found. please make sure the OpsLogix Oracle management packs are imported into your management group."
    return $false
}

#Generate template instance configuration
$NewGUID = [GUID]::NewGuid().ToString().Replace("-","")
$NameSpace = "OpsLogixOracleAlertTemplate_$NewGUID"
$StringBuilder = New-Object System.Text.StringBuilder
$configurationWriter = [System.Xml.XmlWriter]::Create($StringBuilder)
$configurationWriter.WriteStartElement("Configuration");
$configurationWriter.WriteElementString("CounterName", $CounterName);
$configurationWriter.WriteElementString("ColumnName", $ReturnColumnName);
$configurationWriter.WriteElementString("Query", $Query);
$configurationWriter.WriteElementString("IntervalSeconds", $QueryInterval);
$configurationWriter.WriteElementString("HostReference", "");
$configurationWriter.WriteElementString("Target", $TargetClass);
$configurationWriter.WriteElementString("ColumnNamePropertySubs", "`$Data/Property[`@Name='$ReturnColumnName']`$");
$configurationWriter.WriteElementString("Name", $DisplayName);
$configurationWriter.WriteElementString("Description", $Description);
$configurationWriter.WriteElementString("LocaleId", $LocaleId);
$configurationWriter.WriteElementString("ManagementPack", $MPName);
$configurationWriter.WriteElementString("NameSpace", $NameSpace);
$configurationWriter.WriteElementString("Enabled", $Enabled.ToString().ToLower());
$configurationWriter.WriteEndElement();
$configurationWriter.Flush();
$XmlWriter = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackXmlWriter([System.Xml.XmlWriter]::Create($StringBuilder))
$strConfiguration = $StringBuilder.ToString()
Write-Verbose "Template Instance Configuration:"
Write-Verbose $strConfiguration
#Create the template instance
Write-Verbose "Creating the OpsLogix Oracle Performance Collection Rule template instance on management pack '$MPName'..."
Try {
    [Void]$MP.ProcessMonitoringTemplate($OpsLogixOraclePerfTemplate, $strConfiguration, "TemplateoutputOpsLogixIMPOracleConfig2012SQLQueryTemplate$NewGUID", $DisplayName, $Description)
} Catch {
    Write-Error $_.Exception.InnerException
    Return $False
}
#Increase MP version
If ($IncreaseMPVersion)
{
    Write-Verbose "the version of managemnet pack '$MPVersion' will be increased by 0.0.0.1"
    $CurrentVersion = $MP.Version.Tostring()
    $vIncrement = $CurrentVersion.Split('.')
    $vIncrement[$vIncrement.Length - 1] = ([System.Int32]::Parse($vIncrement[$vIncrement.Length - 1]) + 1).ToString()
    $NewVersion = ([System.String]::Join('.', $vIncrement))
    $MP.Version = $NewVersion
}

#Verify and save the template instance
Try {
    $MP.verify()
    $MP.AcceptChanges()
    $bCreated = $true
	Write-Verbose "OpsLogix Oracle Performance Collection Rule template instance '$DisplayName' successfully created in Management Pack '$MPName'($($MP.Version))."
} Catch {
	$MP.RejectChanges()
    $bCreated = $false
    Write-Error "Unable to create OpsLogix Oracle Performance Collection Rule template instance '$DisplayName' in management pack $MPName."
}

#Update sharepoint list
If ($ListItemID -and $ListTitle)
{
    if ($bCreated -eq $true)
    {
        $strResult = "Completed"
    } else {
        "Failed"
    }
    Write-Verbose "Updating SharePoint list item"
    $ListFields = Get-SPListFields -SPConnection $SPConnection -ListName $ListTitle
    $StatusFieldInternalName = ($ListFields | Where-Object {$_.Title -ieq 'Automation Status' -and $_.ReadOnlyField -eq $false}).InternalName
    Write-Verbose "Automation Status Field Internal Name: '$StatusFieldInternalName'"
    $ListFieldValues = @{
        $StatusFieldInternalName = $strResult
    }
    $UpdateSPListItem = Update-SPListItem -SPConnection $SPConnection -ListName $ListTitle -ListItemID $ListItemID -ListFieldsValues $ListFieldValues
}
Write-Output "Done"
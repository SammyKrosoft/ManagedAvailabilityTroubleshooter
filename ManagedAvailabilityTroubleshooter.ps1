<#
.SYNOPSIS
    This script dumps some information that Wade Smith needs to help getting the most of your Exchange / M365 hybrid organization.

.DESCRIPTION
    This script dumps some information that Wade Smith needs to help getting the most of your Exchange / M365 hybrid organization.
    You can dump:
    - Exchange OnPrem related information (general and for Oauth settings checks)
    - Exchange Online related information (same, general and info for Oauth settings checks)
    - MSOL information
    These are necessary to help Wade Smith and his colleagues to help you on challenges you may face on your configuration.

.PARAMETER IncludeUserSpecificInfo
    This parameter is to execute PowerShell collection commands with specific user, domain and org info
    Check and change the variables definitions on the script (will introduce GUI on a later version)

.PARAMETER OnPremExchangeManagementShellCommands
    This is to collect Exchange OnPrem specific information.
    Exchange Management Shell tools are needed to be loaded for this, otherwise
    this will fail data collection.

.PARAMETER OnLineExchangeManagementShellCommands
    This is to collect Exchange Online specific information.
    Exchange Online management module needs to be loaded, otherwise
    this will fail data MS Exchange Online collection.

.PARAMETER MSOLCommands
    This is to collect MS Online (aka Azure) specific information.
    MSOnline module must be loaded, otherwise this will fail MSOL data collection

.INPUTS
    User specific information if you want to use the -IncludeUserSpecificInfo switch

.OUTPUTS
    Many files (see the $OutputFilesCollection Here-String for file names)

.EXAMPLE
Examples to be added later

.EXAMPLE
.\WadeSmithScript.ps1 -CheckVersion
This will dump the script name and current version like :
SCRIPT NAME : WadeSmithScript.ps1
VERSION : v1.0

.NOTES
None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdletBinding(DefaultParameterSetName="NormalRun")]
Param(
    [Parameter(Mandatory = $false, ParameterSetName="NormalRun")][switch]$IncludeUserSpecificInfo,
    [Parameter(Mandatory = $false, ParameterSetName="NormalRun")][switch]$OnPremExchangeManagementShellCommands,
    [Parameter(Mandatory = $false, ParameterSetName="NormalRun")][switch]$OnLineExchangeManagementShellCommands,
    [Parameter(Mandatory = $false, ParameterSetName="NormalRun")][switch]$MSOLCommands,
    [Parameter(Mandatory = $false,ParameterSetName="Check")][switch]$CheckVersion
    
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = "1"
<# Version changes
v1 : added Write-Log and SammyKrosoft Scripting headers
v0.1 : first script version
#>
<<<<<<< HEAD
<# -------------------------- /SCRIPT_HEADER -------------------------- #>

$Script:lastProbeerror = $null
$Script:foundissue = $false
$Script:checkforknownissue =$false
$Script:KnownIssueDetectionAlreadydone = $false
$Script:LoggingMonitoringpath = ""

function TestFileorCmd
{
[cmdletbinding()]
Param( [String] $FileorCmd )

	if ($FileorCmd -like "File missing for this action*")
	{
		Write-Host -foregroundcolor red $FileorCmd;
		exit;
	}
}

function ParseProbeResult
{
[cmdletbinding()]
Param( [String] $FilterXpath , [String] $MonitorToInvestigate , [String] $ResponderToInvestigate)

	TestFileorCmd $ProbeResulteventcmd;
	ParseProbeResult2 ($ProbeResulteventcmd + " -maxevents 200" ) $FilterXpath "Parsing only last 200 probe events for quicker response time" $MonitorToInvestigate $ResponderToInvestigate
	if ("yes","YES","Y","y" -contains (Read-Host ("`nParsed last 200 probe events for quicker response.`nDo you like to parse all probe events ? Y/N (default is ""N"")")))
	{ ParseProbeResult2 $ProbeResulteventcmd $FilterXpath "Parsing all probe events. this may be slow as there is lots of events" $MonitorToInvestigate $ResponderToInvestigate}
}

function ParseProbeResult2
{
[cmdletbinding()]
Param( [String] $ProbeResulteventcompletecmd , [String] $FilterXpath , [String] $waitstring , [String] $MonitorToInvestigate , [String] $ResponderToInvestigate)

	TestFileorCmd $ProbeResulteventcmd;
	$Probeeventscmd = '(' + $ProbeResulteventcompletecmd + ' -FilterXPath ("' + $FilterXpath +'") -ErrorAction SilentlyContinue | % {[XML]$_.toXml()}).event.userData.eventXml'	
	Write-verbose $Probeeventscmd
	$titleprobeevents = "Probe events"
	if ( $ProbeDetailsfullname )
	{	$titleprobeevents = $ProbeDetailsfullname + " events" }	
	if ($waitstring)
	{
		write-progress "Checking Probe Result Events" -status $waitstring
	}
	else
	{
		write-progress "Checking Probe Result Events"
	}
	$checkerrorcount = $error.count
	$Probeevents = invoke-expression $Probeeventscmd
	write-progress "Checking Probe Result Events" -completed
	$checkerrorcount = $error.count - $checkerrorcount
	if ($checkerrorcount -gt 0)
	{
		for ($j=0;$j -lt $checkerrorcount;$j++)
		{
			if ($error[$j].FullyQualifiedErrorId -like "NoMatchingEventsFound*")
			{ write-host -foreground red "No events were found"}
			else
			{ write-host -foreground red $error[$j].exception.message}
		}
	}
	if ($Probeevents)
	{
		foreach ($Probeevt in $Probeevents)
		{
		    If ($Probeevt.ResultType -eq 4)
		    {
			$Script:lastProbeerror = $Probeevt
			if ($Script:KnownIssueDetectionAlreadydone -eq $false) {KnownIssueDetection $MonitorToInvestigate $ResponderToInvestigate}
			Break;
		    }
		}
		if ($Script:KnownIssueDetectionAlreadydone -eq $false) {KnownIssueDetection $MonitorToInvestigate $ResponderToInvestigate}
		If (!($ExportCSVInsteadOfGridView)){
			$Probeevents | Select-Object -Property @{n="ExecutionStartTime (GMT)";e={$_.ExecutionStartTime}},@{n="ExecutionEndTime (GMT)";e={$_.ExecutionEndTime}},@{n='ResultType';e={$_.ResultType -replace "1","Timeout"-replace "2","Poisoned" -replace "3","Succeeded" -replace "4","Failed" -replace "5","Quarantined" -replace "6","Rejected"}},@{n='Error';e={$_.Error -replace "`r`n","`r"}},@{n='Exception';e={$_.Exception -replace "`r`n","`r"}},FailureContext,@{n='ExecutionContext';e={$_.ExecutionContext -replace "`r`n","`r"}},RetryCount,ServiceName,ResultName,StateAttribute*| Out-GridView -title $titleprobeevents
		} Else {
			$Probeevents | Select-Object -Property @{n="ExecutionStartTime (GMT)";e={$_.ExecutionStartTime}},@{n="ExecutionEndTime (GMT)";e={$_.ExecutionEndTime}},@{n='ResultType';e={$_.ResultType -replace "1","Timeout"-replace "2","Poisoned" -replace "3","Succeeded" -replace "4","Failed" -replace "5","Quarantined" -replace "6","Rejected"}},@{n='Error';e={$_.Error -replace "`r`n","`r"}},@{n='Exception';e={$_.Exception -replace "`r`n","`r"}},FailureContext,@{n='ExecutionContext';e={$_.ExecutionContext -replace "`r`n","`r"}},RetryCount,ServiceName,ResultName,StateAttribute*| Export-CSV -NoTypeInformation -Path "$PSScriptRoot\ProbeEvents_$(Get-Date -Format yyyyMMdd-HHmmss).csv"
		}
	}
	if ($Script:KnownIssueDetectionAlreadydone -eq $false) {KnownIssueDetection $MonitorToInvestigate $ResponderToInvestigate}
}


function InvestigateProbe
{
[cmdletbinding()]
Param([String]$ProbeToInvestigate , [String]$MonitorToInvestigate , [String]$ResponderToInvestigate , [String]$ResourceNameToInvestigate , [String]$ResponderTargetResource )

	TestFileorCmd $ProbeDefinitioneventcmd;
    if (-Not ($ResponderTargetResource) -and ($ProbeToInvestigate.split("/").Count -gt 1))
    {
        $ResponderTargetResource = $ProbeToInvestigate.split("/")[1]
    }
	$ProbeDetailscmd = '(' + $ProbeDefinitioneventcmd + '| % {[XML]$_.toXml()}).event.userData.eventXml| ? {$_.Name -like "' + $ProbeToInvestigate.split("/")[0] + '*" }'
	Write-verbose $ProbeDetailscmd
	write-progress "Checking Probe definition"
	$ProbeDetails = invoke-expression $ProbeDetailscmd
	write-progress "Checking Probe definition" -completed
	if ( $ProbeDetails)
	{
		if ($ProbeDetails.Count -gt 1)
		{
			if ($ResourceNameToInvestigate)
			{
				$ProbeDetailsforselectedResourceName = $ProbeDetails | Where-Object {$_.TargetResource -eq $ResourceNameToInvestigate}
				if ($ProbeDetailsforselectedResourceName )
				{   $ProbeDetails = $ProbeDetailsforselectedResourceName }
			}
			if ($ProbeDetails.Count -gt 1)
			{
				if ($ResponderTargetResource)
				{
					$ProbeDetailsforselectedResourceName = $ProbeDetails | Where-Object {$_.TargetResource -eq $ResponderTargetResource}
					if ($ProbeDetailsforselectedResourceName )
					{   $ProbeDetails = $ProbeDetailsforselectedResourceName }
				}

				if ($ProbeDetails.Count -gt 1)
				{
					Write-Host -foregroundcolor red ("Found no probe for " + $ResourceNameToInvestigate + " TargetResource")
					Write-Host "`nSelected all possible Probes in this list: "
					if ($ProbeDetails.Count -gt 20)
					{
						Write-host -foregroundcolor red ("more than 30 Probes in the list. Keeping only the 30 first probes")
						$ProbeDetails = $ProbeDetails[0..19]
					}
				}
			}
		}
		$ProbeDetails | Format-List *
		$ProbeDetailsfullname = $null								
		foreach ($ProbeInfo in $ProbeDetails)
		{
			$probename2add = $ProbeInfo.Name
			if ($ProbeInfo.TargetResource)
			{
				if ( -not ($ProbeInfo.TargetResource -eq "[null]"))
				{ $probename2add += "/" + $ProbeInfo.TargetResource}
			}
			if ($ProbeDetailsfullname -eq $null )
			{$ProbeDetailsfullname = $Probename = $probename2add }
			else
			{
				$ProbeNameAlreadyinthelist = $false
				foreach ( $PresentProbeName in ($Probename -replace " and ",";").split(";"))
				{
					if ($PresentProbeName -eq $probename2add)
					{$ProbeNameAlreadyinthelist = $true}
				}
				if ($ProbeNameAlreadyinthelist -eq $false)
				{
					$ProbeDetailsfullname += "' or ResultName='" + $probename2add
					$Probename += " and " + $probename2add
				}
			}
		}
							
		if ($MonitorToInvestigate)
		{
			$relationdescription = "`n" + $Probename +" errors can result in the failure of " + $MonitorToInvestigate + " monitor"
			if ( $ResponderToInvestigate)
			{
				$relationdescription +=	" which triggered " + $ResponderToInvestigate
			}
			Write-host $relationdescription
		}
		If ( $Probename -eq "EacBackEndLogonProbe")
		{
			if ($Script:KnownIssueDetectionAlreadydone -eq $false) {KnownIssueDetection $MonitorToInvestigate $ResponderToInvestigate}
			
			$EacBackEndLogonProbefolder = $Script:LoggingMonitoringpath +"\ECP\EacBackEndLogonProbe"
			if ( Test-Path $EacBackEndLogonProbefolder)
			{
				$EacBackEndLogonProbefile = Get-ChildItem ($EacBackEndLogonProbefolder) | Select-Object -last 1
				if ($EacBackEndLogonProbefile)
				{
					write-host "found and opening EacBackEndLogonProbe log / check the file for further error details"
					notepad $EacBackEndLogonProbefile.fullname
				}
			}
			else
			{ write-host -foregroundcolor red ("Missing logs from path $EacBackEndLogonProbefolder ")}
		}
		else
		{
			ParseProbeResult ("*[UserData[EventXML[ResultName='" + $ProbeDetailsfullname + "']]]") $MonitorToInvestigate $ResponderToInvestigate
		}
	}
	else
	{   write-host("`nFound no definitions for " + $ProbeToInvestigate + " probe") }
}

Function InvestigateMonitor
=======
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
# Log or report file definition
$UserDocumentsFolder = "$($env:Userprofile)\Documents"
$OutputReport = "$UserDocumentsFolder\$($ScriptName)_Output_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$UserDocumentsFolder\$($ScriptName)_Logging_$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
function Write-Log
>>>>>>> 6d9361324d596d39dbf56670daab092c59007f88
{
	<#
	.SYNOPSIS
		This function creates or appends a line to a log file.
	.PARAMETER  Message
		The message parameter is the log message you'd like to record to the log file.
	.EXAMPLE
		PS C:\> Write-Log -Message 'Value1'
		This example shows how to call the Write-Log function with named parameters.
	#>
	[CmdletBinding()]
	param (
		[Parameter(Mandatory=$true,position = 0)]
		[string]$Message,
		[Parameter(Mandatory=$false,position = 1)]
        [string]$LogFileName=$ScriptLog,
        [Parameter(Mandatory=$false, position = 2)][switch]$Silent
	)
	
	try
	{
		$DateTime = Get-Date -Format 'MM-dd-yy HH:mm:ss'
		$Invocation = "$($MyInvocation.MyCommand.Source | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)"
		Add-Content -Value "$DateTime - $Invocation - $Message" -Path $LogFileName
		if (!($Silent)){Write-Host $Message -ForegroundColor Green}
	}
	catch
	{
		Write-Error $_.Exception.Message
	}
}
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
Write-Log "************************** Script Start **************************"

#Collect PowerShell command result in txt files: 

# Variables declaration
$OutputFilesCollection = @"
$($env:Userprofile)\Documents\OnPrem_OrgRel.txt
$($env:Userprofile)\Documents\OnPrem_Test-OrganizationRelationship.txt
$($env:Userprofile)\Documents\OnPrem_IntraOrgCon.txt
$($env:Userprofile)\Documents\OnPrem_AvaiAddSpa.txt
$($env:Userprofile)\Documents\OnPrem_SharingPolicy.txt
$($env:Userprofile)\Documents\OnPrem_WebSerVirDir.txt
$($env:Userprofile)\Documents\OnPrem_AutoDVirDir.txt
$($env:Userprofile)\Documents\OnPrem_FedTrust.txt
$($env:Userprofile)\Documents\OnPrem_FedOrgIden.txt
$($env:Userprofile)\Documents\OnPrem_FedInfo.txt
$($env:Userprofile)\Documents\OnPrem_TestFedTrust.txt
$($env:Userprofile)\Documents\OnPrem_TestFedCert.txt
$($env:Userprofile)\Documents\OnPrem_RemoteMailbox.txt
$($env:Userprofile)\Documents\OnPrem_Mailbox.txt
$($env:Userprofile)\Documents\OnPrem_Server.txt
$($env:Userprofile)\Documents\OnPrem_HybridConfig.txt
$($env:Userprofile)\Documents\O365_OrgRel.txt
$($env:Userprofile)\Documents\O365_Test-OrganizationRelationship.txt
$($env:Userprofile)\Documents\O365_IntraOrgCon.txt
$($env:Userprofile)\Documents\O365_AvaiAddSpa.txt
$($env:Userprofile)\Documents\O365_SharingPolicy.txt
$($env:Userprofile)\Documents\O365_FedTrust.txt
$($env:Userprofile)\Documents\O365_FedInfo.txt
$($env:Userprofile)\Documents\O365_FedOrgIden.txt
$($env:Userprofile)\Documents\O365_MailUser.txt
$($env:Userprofile)\Documents\O365_Mailbox.txt
$($env:Userprofile)\Documents\OnPrem_IntraOrgCon.txt
$($env:Userprofile)\Documents\OnPrem_IntraOrgConfig.txt
$($env:Userprofile)\Documents\OnPrem_AuthServer.txt
$($env:Userprofile)\Documents\OnPrem_ParApp.txt
$($env:Userprofile)\Documents\OnPrem_PartnerAppAcct.txt
$($env:Userprofile)\Documents\OnPrem_AuthConfig.txt
$($env:Userprofile)\Documents\OnPrem_AuthConfigCert.txt
$($env:Userprofile)\Documents\OnPrem_WebSerVirDir.txt
$($env:Userprofile)\Documents\OnPrem_AutoDVirDir.tx
$($env:Userprofile)\Documents\OnPrem_OrgRel.txt
$($env:Userprofile)\Documents\OnPrem_AvaiAddSpa.txt
$($env:Userprofile)\Documents\OnPrem_TestOAuthConnectivityEWS.txt
$($env:Userprofile)\Documents\OnPrem_TestOAuthConnectivityAutoD.txt
$($env:Userprofile)\Documents\OnPrem_RemoteMailbox.txt
$($env:Userprofile)\Documents\OnPrem_Mailbox.txt
$($env:Userprofile)\Documents\OnPrem_Server.txt
$($env:Userprofile)\Documents\OnPrem_ExchangeCertificates.txt
$($env:Userprofile)\Documents\OnPrem_HybridConfig.txt
$($env:Userprofile)\Documents\O365_IntraOrgCon.txt
$($env:Userprofile)\Documents\O365_IntraOrgConfig.txt
$($env:Userprofile)\Documents\O365_AuthServer.txt
$($env:Userprofile)\Documents\O365_PartnerApp.txt
$($env:Userprofile)\Documents\O365_TestOAuthConnectivityEWS.txt
$($env:Userprofile)\Documents\O365_TestOAuthConnectivityAutoD.txt
$($env:Userprofile)\Documents\O365_OrgRel.txt
$($env:Userprofile)\Documents\O365_MailUser.txt
$($env:Userprofile)\Documents\O365_Mailbox.txt
$($env:Userprofile)\Documents\Msol_ServicePrincipal.txt
$($env:Userprofile)\Documents\Msol_ServicePrincipalNames.txt
$($env:Userprofile)\Documents\Msol_ServicePrincipalCredential.txt
"@ -split "`n" | ForEach-Object { $_.trim() }

If ($IncludeUserSpecificInfo){
    Write-Log "Including user specific information..."
    
    $OnPremisesMailbox = "User1@Contoso.ca"
    $CloudMailbox = "UserCloud1@Contoso.ca"
    $CustomerOnMicrosoftDomain = "Contoso.mail.onmicrosoft.com"
    $CustomerDomain = "Contoso.ca"
    $OnPremisesExternalEWSURL = "https://mail.domain.com/ews/exchange.asmx"
    $OnPremisesAutodiscoverURL = "https://mail.domain.com/autodiscover/autodiscover.xml"
    
    Write-Log "OnPrem Mailbox: $OnPremisesMailbox"
    Write-Log "Cloud Mailbox: $CloudMailbox"
    Write-Log "Customer OnMicrosoft Domain : $CustomerOnMicrosoftDomain"
    Write-Log "Curstomer Domain: $CustomerDomain"
    Write-Log "On-Premises External EWS URL: $OnPremisesExternalEWSURL"
    Write-Log "On-Premises Autodiscover URL: $OnPremisesAutodiscoverURL"
}

# -------------------------------------------------------------------------------------------------
# In Exchange On-premises<Connect to Exchange management Shell>
# -------------------------------------------------------------------------------------------------
If ($OnPremExchangeManagementShellCommands){
    Write-Log "Used -OnPremExchangeManagementShellCommands switch ... dumping Exchange OnPrem info"
    Get-FederationTrust | Set-FederationTrust -RefreshMetadata 
    Get-AutoDiscoverVirtualDirectory | FL > $OutputFilesCollection[6]
    Get-AvailabilityAddressSpace | FL > $OutputFilesCollection[3]
    Get-FederatedOrganizationIdentifier | FL > $OutputFilesCollection[8]
    Get-FederationTrust | FL > $OutputFilesCollection[7]
    Get-HybridConfiguration | FL > $OutputFilesCollection[15]
    Get-IntraOrganizationConnector | FL > $OutputFilesCollection[2]
    Get-OrganizationRelationship | FL > $OutputFilesCollection[0]
    Get-ExchangeServer | FT name, serverrole, AdminDisplayVersion > $OutputFilesCollection[14]
    Get-SharingPolicy | FL > $OutputFilesCollection[4]
    Test-FederationTrustCertificate | FL > $OutputFilesCollection[11]
    Get-WebServicesVirtualDirectory | FL > $OutputFilesCollection[5]

    If ($IncludeUserSpecificInfo){
        Write-Log "Used -IncludeUserSpecificInfo switch ... dumping User specific info for Exchange OnPrem"
        # User specific informtion
        Get-FederationInformation -Domainname $CustomerOnMicrosoftDomain | FL > $OutputFilesCollection[9]
        # User specific information
        Get-Mailbox $OnPremisesMailbox | FL > $OutputFilesCollection[13]
        # User specific information
        Get-RemoteMailbox $CloudMailbox | FL > $OutputFilesCollection[12]
        # User specific information
        Test-FederationTrust -USerIdentity $OnPremisesMailbox > $OutputFilesCollection[10]
        Test-OrganizationRelationship -Identity "On-premises to O365 Organization Relationship" -UserIdentity $OnPremisesMailbox -Verbose > $OutputFilesCollection[1]
    }
}

# -------------------------------------------------------------------------------------------------
# In Exchange Online<Connect to Exchange Online service>ï¼š  
# -------------------------------------------------------------------------------------------------
If ($OnLineExchangeManagementShellCommands){
    Write-Log "Used -OnLineExchangeManagementShellCommands switch ... dumping Exchange OnLine info"
    Get-AvailabilityAddressSpace |  FL > $OutputFilesCollection[19]
    Get-FederatedOrganizationIdentifier | FL > $OutputFilesCollection[23]
    Get-FederationTrust | FL > $OutputFilesCollection[21]
    Get-IntraOrganizationConnector | FL > $OutputFilesCollection[18]
    Get-OrganizationRelationship | FL > $OutputFilesCollection[16]
    Get-SharingPolicy | FL > $OutputFilesCollection[20]
    If ($IncludeUserSpecificInfo){
        Write-Log "Used -IncludeUserSpecificInfo switch ... dumping User specific info for Exchange Online"
        # User specific information
        Get-FederationInformation -DomainName $CustomerDomain | FL > $OutputFilesCollection[22]
        # User specific information
        Get-Mailbox $CloudMailbox | FL > $OutputFilesCollection[25]
        Get-MailUser $OnPremisesMailbox | FL  > $OutputFilesCollection[24]
        # User specific information
        Test-OrganizationRelationship -UserIdentity $CloudMailbox  -Identity "Exchange Online to On Premises Organization Relationship" -Verbose > $OutputFilesCollection[17]
    }
}

# -------------------------------------------------------------------------------------------------
#We also need to check the Oauth settings: 
# In Exchange On-premises<Connect to Exchange management Shell>ï¼š 
# -------------------------------------------------------------------------------------------------
If ($OnPremExchangeManagementShellCommands){
    Write-Log "Now dumping Oauth related information."
    Write-Log "Used -OnPremExchangeManagementShellCommands switch ... dumping Exchange OnPrem info for Oauth settings"
    Get-AuthConfig | FL > $OutputFilesCollection[31]
    Get-ExchangeCertificate -Thumbprint (Get-AuthConfig).CurrentCertificateThumbprint | FL > $OutputFilesCollection[32]
    Get-AuthServer | FL > $OutputFilesCollection[28]
    Foreach ($i in (Get-ExchangeServer)) {Write-Host $i.FQDN; Get-ExchangeCertificate -Server $i.Identity} > $OutputFilesCollection[42]
    Get-IntraOrganizationConfiguration | FL > $OutputFilesCollection[27]
    Get-PartnerApplication | FL > $OutputFilesCollection[29]
    Get-PartnerApplication 00000002-0000-0ff1-ce00-000000000000 | Select-Object -ExpandProperty LinkedAccount | Get-User | FL > $OutputFilesCollection[30]
    If ($IncludeUserSpecificInfo){
        Write-Log "Used -IncludeUserSpecificInfo switch ... dumping User specific info for Exchange OnPrem for Oauth settings"
        # User specific information
        Test-OAuthConnectivity -Service AutoD  -TargetUri https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc -Mailbox $OnPremisesMailbox -Verbose | FL > $OutputFilesCollection[38]
        Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/ews/exchange.asmx -Mailbox $OnPremisesMailbox -Verbose | FL > $OutputFilesCollection[37]
    }
}

# -------------------------------------------------------------------------------------------------
# In Exchange Online<Connect to Exchange Online service>ï¼š 
# -------------------------------------------------------------------------------------------------
If ($OnLineExchangeManagementShellCommands){
    Write-Log "Now dumping Oauth related information."
    Write-Log "Used -OnLineExchangeManagementShellCommands switch ... dumping Exchange OnLine info for Oauth settings"
    if ($IncludeUserSpecificInfo){
        Write-Log "Used -IncludeUserSpecificInfo switch ... dumping User specific info for Exchange OnLine for Oauth settings"
        # User specific information
        Test-OAuthConnectivity -Service AutoD -TargetUri $OnPremisesAutodiscoverURL -Mailbox $CloudMailbox -Verbose | FL > $OutputFilesCollection[49]
        Test-OAuthConnectivity -Service EWS -TargetUri $OnPremisesExternalEWSURL -Mailbox $CloudMailbox -Verbose | FL > $OutputFilesCollection[48]
    }
    Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | FL > $OutputFilesCollection[46]
    Get-IntraOrganizationConfiguration | FL > $OutputFilesCollection[45]
    Get-PartnerApplication | FL > $OutputFilesCollection[47]
}
# -------------------------------------------------------------------------------------------------
# Azure/MSOLPowershell: 
# -------------------------------------------------------------------------------------------------
If ($MSOLCommands){
    Write-Log "Used -MSOLCommands switch ... dumping MS OnLine Azure info for Oauth settings"
    Get-MsolServicePrincipal -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" | FL  > $OutputFilesCollection[53]
    (Get-MsolServicePrincipal -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000").ServicePrincipalNames > $OutputFilesCollection[54]
    Get-MsolServicePrincipalCredential -ServicePrincipalName "00000002-0000-0ff1-ce00-000000000000" -ReturnKeyValues $true > $OutputFilesCollection[55]
}


<<<<<<< HEAD
	$CheckRecoveryActionForMultipleMachines = $RecoveryActionResultscmd -like "*Foreach-Object*"
	$RecoveryActions = $null
	if ($CheckRecoveryActionForMultipleMachines)
	{ TestFileorCmd ($RecoveryActionResultscmd + ")}")}
	else
	{ TestFileorCmd ($RecoveryActionResultscmd + ")") }
	$RecoveryActionscmd = $RecoveryActionResultscmd + '| % {[XML]$_.toXml()}).event.userData.eventXml'
	if ($Investigationchoose -eq 0)
	{ $RecoveryActionscmd += '| ? {$_.Id -eq "ForceReboot"}' }
	if ($CheckRecoveryActionForMultipleMachines)
	{ $RecoveryActionscmd += '; For ($i=$RAindex; $i -lt $RecoveryActions.Count; $i++) { $RecoveryActions[$i]|Add-Member -Name "MachineName" -Value $exserver -MemberType NoteProperty}};$RecoveryActions' }
	Write-verbose $RecoveryActionscmd
	write-progress "Checking Recovery Actions"
	$RecoveryActions = invoke-expression $RecoveryActionscmd
	write-progress "Checking Recovery Actions" -completed
	if ($RecoveryActions)
	{	
		if ($Investigationchoose -eq 0)
		{
			write-host ("`nLast Reboot was triggered by the Responder "+ $RecoveryActions[0].RequestorName + " at " + $RecoveryActions[0].StartTime + " ." )
			$SelectTitle = "Select the ForceReboot that you like to investigate"
		}
		else
		{   $SelectTitle = "Select the Recovery Action that you like to investigate" }
		write-host $SelectTitle
		Start-Sleep -s 1
		$RAoutgridviewcmd = '$RecoveryActions | select -Property '
		if ($CheckRecoveryActionForMultipleMachines)
		{$RAoutgridviewcmd += "MachineName,"}
		$RAoutgridviewcmd += '@{n="StartTime (GMT)";e={$_.StartTime}}, @{n="EndTime (GMT)";e={$_.EndTime}} , Id , ResourceName , InstanceId , RequestorName , Result , State , ExceptionName,ExceptionMessage,LamProcessStartTime,ThrottleIdentity , ThrottleParametersXml , Context | Sort-Object "StartTime (GMT)" -Descending | Out-GridView -PassThru -title $SelectTitle'
		Write-verbose $RAoutgridviewcmd
		$RecoveryActionToInvestigate = invoke-expression $RAoutgridviewcmd
		if ($RecoveryActionToInvestigate)
		{
			if ($RecoveryActionToInvestigate.Count -gt 1 )
			{   $RecoveryActionToInvestigate = $RecoveryActionToInvestigate[0] }
			if ($CheckRecoveryActionForMultipleMachines)
			{
				if ([string]::Compare($RecoveryActionToInvestigate.MachineName,$env:computername,$true) -ne 0)
				{
					Write-host -foregroundcolor yellow ("`nThe RecoveryAction you select is regarding a different server : " + $RecoveryActionToInvestigate.MachineName + " .")
					Write-host -foregroundcolor yellow ("Run this script on this server directly to analyse this RecoveryAction further." )
					exit;
				}
			}
			InvestigateResponder $RecoveryActionToInvestigate.RequestorName $RecoveryActionToInvestigate.ResourceName
		}
		else
		{   if ($Investigationchoose -eq 0) {Write-host ("`nYou have not selected any occurrence. Run the script again and select an occurrence" ) }}    
	}
	else
	{   write-host "`nFound no event with ID ForceReboot in RecoveryActionResults log. Health Manager shouldn't have triggered a reboot recently." }
}

if ($Investigationchoose -eq 2)
{
	$SpecificResponderorMonitororProbe = Read-Host ("Enter the name of the Responder/Monitor or Probe ")
	if ($SpecificResponderorMonitororProbe)
	{
		$IsitaResponderorMonitororProbe = 0
		if ($SpecificResponderorMonitororProbe.split("/")[0].ToLower().EndsWith("probe"))
		{		
			$IsitaResponderorMonitororProbe = 2
		}
		elseif ($SpecificResponderorMonitororProbe.split("/")[0].ToLower().EndsWith("monitor"))
		{
			$IsitaResponderorMonitororProbe = 1
		}
		else
		{
			$IsResponder = New-Object System.Management.Automation.Host.ChoiceDescription "&Responder", "Responder"
			$IsMonitor = New-Object System.Management.Automation.Host.ChoiceDescription "&Monitor", "Monitor"
			$IsProbe = New-Object System.Management.Automation.Host.ChoiceDescription "&Probe", "Probe"
			$IsitaResponderorMonitororProbe = $host.ui.PromptForChoice("", "Is it a : ", [System.Management.Automation.Host.ChoiceDescription[]]($IsResponder, $IsMonitor,$IsProbe), 0)
		}
		switch ( $IsitaResponderorMonitororProbe)
		{
			0 { InvestigateResponder $SpecificResponderorMonitororProbe $null}
			1 { InvestigateMonitor $SpecificResponderorMonitororProbe $null $null $null }
			2 { InvestigateProbe $SpecificResponderorMonitororProbe $null $null $null $null }
		}
	}
	else
	{ write-host -foregroundcolor red ("No name specified")}
	exit
}
if ($Investigationchoose -eq 3)
{
	$CheckAlertsForMultipleMachines = $ManagedAvailabilityMonitoringcmd -like "*Foreach-Object*"
	$alertevents =$null
	if ($CheckAlertsForMultipleMachines)
	{TestFileorCmd ($ManagedAvailabilityMonitoringcmd + " }") }
	else
	{TestFileorCmd $ManagedAvailabilityMonitoringcmd }
	$ManagedAvailabilityMonitoringcmd = $ManagedAvailabilityMonitoringcmd  + '-maxevents 200 |? {$_.Id -eq 4 }'
	if ($CheckAlertsForMultipleMachines)
	{$ManagedAvailabilityMonitoringcmd += ' };$alertevents'}
	Write-verbose $ManagedAvailabilityMonitoringcmd
	write-progress "Checking SCOM Alerts"
	$alertevents = invoke-expression $ManagedAvailabilityMonitoringcmd 
	$alerteventsprops = ($alertevents | ForEach-Object {[XML]$_.toXml()}).event.userData.eventXml
	For ($i=0; $i -lt $alerteventsprops.Count; $i++) 
	{
		$alerteventsprops[$i] | Add-Member TimeCreated $alertevents[$i].TimeCreated
		if ($CheckAlertsForMultipleMachines)
		{ $alerteventsprops[$i] | Add-Member MachineName $alertevents[$i].MachineName }
	}
	write-progress "Checking SCOM Alerts" -completed
	$alertoutgridviewcmd = '$alerteventsprops | select -Property '
	if ($CheckAlertsForMultipleMachines)
	{$alertoutgridviewcmd += "MachineName," }
	
	if (!($ExportCSVInsteadOfGridView)){
		$alertoutgridviewcmd += 'TimeCreated, Monitor,HealthSet,Subject,Message | Out-GridView -title "SCOM Alerts"'
	} Else {
		$alertoutgridviewcmd += 'TimeCreated, Monitor,HealthSet,Subject,Message | Export-CSV -NoTypeInformation -Path "$PSScriptRoot\SCOMAlerts_$(Get-Date -Format yyyyMMdd-HHmmss).csv"'
	}
    Write-host "Running $alertoutgridviewcmd"
	invoke-expression $alertoutgridviewcmd
}
if ($Investigationchoose -eq 4)
{
	InvestigateUnhealthyMonitor $ServerHealthfile
}
if ($Investigationchoose -eq 5)
{	
	ParseProbeResult "*[UserData[EventXML [ResultType='4']]]" $null $null
}
if (($Investigationchoose -eq 6) -and ($exchangeversion))
{	
	CollectMaLogs $MyInvocation.MyCommand.Path
}
=======
<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>
>>>>>>> 6d9361324d596d39dbf56670daab092c59007f88

<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
Write-Log "************************** Script End **************************"
$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
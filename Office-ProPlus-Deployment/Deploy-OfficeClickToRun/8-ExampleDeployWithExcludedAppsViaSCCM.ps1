#  Office ProPlus Click-To-Run Deployment Script example
#
#  This script demonstrates how utilize the scripts in OfficeDev/Office-IT-Pro-Deployment-Scripts repository together to create
#  Office ProPlus Click-To-Run deployment script that will be adaptive to the configuration of the computer it is run from

Process {
 $scriptPath = "."

 if ($PSScriptRoot) {
   $scriptPath = $PSScriptRoot
 } else {
   $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
 }

#Importing all required functions
. $scriptPath\Generate-ODTConfigurationXML.ps1
. $scriptPath\Edit-OfficeConfigurationFile.ps1
. $scriptPath\Install-OfficeClickToRun.ps1
. $scriptPath\Setup-SCCMOfficeDeployment.ps1

$targetFilePath = "Configuration_UpdateSource.xml"

#This example will create an Office Deployment Tool (ODT) configuration file and include all of the Languages currently in use on the computer
#from which the script is run.  It will then remove the Version attribute from the XML to ensure the installation gets the latest version and 
#will set the branch to 'Business'.  It will then detect if O365ProPlusRetail or O365BusinessRetail is in the configuration file and if so
#it will add Lync and Groove to the excluded apps. It will then initiate a install

Generate-ODTConfigurationXml -Languages AllInUseLanguages -TargetFilePath $targetFilePath | Set-ODTAdd -Version $NULL -Branch Business | Out-Null

if ((Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId O365ProPlusRetail)) {
     Set-ODTProductToAdd -ProductId "O365ProPlusRetail" -TargetFilePath $targetFilePath -ExcludeApps ("Lync", "Groove") | Out-Null
}

if ((Get-ODTProductToAdd -TargetFilePath $targetFilePath -ProductId O365BusinessRetail)) {
     Set-ODTProductToAdd -ProductId "O365BusinessRetail" -TargetFilePath $targetFilePath -ExcludeApps ("Lync", "Groove") | Out-Null
}

#Download content for distribution and install
Download-OfficeFiles -Path \\SCCM-CM\OfficeDeployment2

#setup SCCM package
Setup-SCCMOfficeProPlusPackage -Path \\SCCM-CM\OfficeDeployment2 -PackageName "Office ProPlus Deployment test2" -ProgramName "Office2016Setup.exe" -distributionPoint SCCM-CM.CONTOSO.COM


#deploy package -note: only to be used after the package has completed distribution
Deploy-SCCMOfficeProPlusPackage -Collection "collection test 21" -PackageName "Office ProPlus Deployment test2" -ProgramName Office2016Setup.exe

}
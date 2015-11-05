﻿<#
.SYNOPSIS
Determines the size of the update and quality of delta compression
for an office update.

.DESCRIPTION
Uses Office Deployment Tool to download and install a specified starting
version of Office. Then captures the current received bytes on the 
NetAdapter, before starting an update to the specified end version.
Records the end received bytes to determine the total size of the 
download. Then zips the apply folder within the office updates folder
to determine what the max download size would be without delta
compression. Comparing these two values provides the delta compression
value.

.Notes
Recommended use is running this script on a clean VM.

.PARAMETER VersionStart
The version to install initially before updating

.PARAMETER VersionEnd
The version to update to

.Example
./Check-OfficeUpdateNetworkLoad -VersionStart 15.0.4623.1003 -VersionEnd 15.0.4631.1002
Installs Version 15.04623.1003 and updates to version 15.0.4631.1002 and returns the 
network traffic numbers. (In original test environment this call returned the values
MaxDownload: ~324000000, ActualDownload: ~128500000, DeltaCompressionRate: ~0.60)

.Outputs
Hashtable with values for Downloaded bytes, max size, delta compression rate

#>

Param(
    [Parameter(Mandatory=$true)]
    [string] $VersionStart,

    [Parameter(Mandatory=$true)]
    [string] $VersionEnd

)

Begin{
$ZipPath = "$env:USERPROFILE\Downloads\sizeTest.zip"
$config1 = 
"<Configuration>
    <Add OfficeClientEdition=`"32`" Version=`"$VersionStart`">
        <Product ID=`"O365ProPlusRetail`">
            <Language ID=`"en-us`" />
        </Product>
    </Add>
    <Updates Enabled=`"TRUE`" />
</Configuration>  "
$folderPath = "$env:ProgramFiles\Microsoft Office 15\Data\Updates\Apply"
$ODTSource = "http://download.microsoft.com/download/6/2/3/6230F7A2-D8A9-478B-AC5C-57091B632FCF/officedeploymenttool_x86_4747-1000.exe"
}

Process{
#download setup
Invoke-WebRequest $ODTSource -OutFile "$env:USERPROFILE\Downloads\officedeploymenttool_x86_4747-1000.exe" | Out-Null
Set-Location "$env:USERPROFILE\Downloads"
.\officedeploymenttool_x86_4747-1000.exe /extract:$env:USERPROFILE\downloads\ODT /passive /quiet | Out-Null
Set-Location ODT

#build configuration file
$config1 | Out-File configuration.xml
./setup.exe /configure configuration.xml | Out-Null

#Start word to block update from applying when finished downloading
Start-Process "${env:ProgramFiles}\Microsoft Office 15\root\office15\WINWORD.EXE"

#get bytes for net adapter
$netstat1 = Get-NetAdapterStatistics

#Start update
Start-Process "${env:ProgramFiles}\Microsoft Office 15\Clientx64\OfficeC2RClient.exe" "/update user updatetoversion=$VersionEnd"

#Wait for update to complete and stop the UAC process if it gets in the way
$complete = $false
while($complete -eq $false){
    $procs = Get-Process | ? ProcessName -eq 'officeclicktorun'
    $UACProc = Get-Process | ? ProcessName -eq "consent"
    if($UACProc -ne $null){
        $UACProc.Kill()
        $UACProc = $null
        $complete = $true
    }
    foreach($proc in $procs){
        if($proc.MainWindowTitle -eq "Please close programs" -or $proc.MainWindowTitle -eq "We need to close some programs"){
            $complete = $true
        }
    }
}

#get bytes for net adapter
$netstat2 = Get-NetAdapterStatistics
$bytes = 0
if($netstat1.GetType() -is [array]){
    foreach($adapter in $netstat2){
        $bytes += $adapter.ReceivedBytes
    }
    foreach($adapter in $netstat1){
        $bytes -= $adapter.ReceivedBytes
    }
}else{
    $bytes = $netstat2.ReceivedBytes - $netstat1.ReceivedBytes 
}

#Zip the Data/Updates/Apply folder to get what size of update could have been
Add-Type -assembly "system.io.compression.filesystem"
[io.compression.zipfile]::CreateFromDirectory($folderPath, $ZipPath)
$zipSize = Get-Item $ZipPath

#Stop word process
$word = Get-Process | ? ProcessName -eq WINWORD
$word.Kill()
$word = $null

#Output results
@{
    ActualDownload = $bytes;
    MaxDownload = $zipSize.Length;
    DeltaCompressionRate = 1 - ($bytes/$zipSize.Length);
}
}
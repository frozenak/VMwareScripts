<#
.SYNOPSIS
    HTML report on snapshots and disk consolidation for multiple vCenter environments.
.DESCRIPTION
    Basis of script: https://community.spiceworks.com/scripts/show/1871-vm-snapshot-report
    Function Set-AlternatingRows to alternate row colors.
    Function Get-SnapshotCreator to find user that created VM snapshot.
    Function Script_Logging to log actions
    Runs a for each loop on list of vCenters to gather VM snapshots, VM snapshot creator, and find any VM disks that may need consolidation.

    v2 - Created foreach loop, replace text to change font color, add bold.
    v2a - Cleaned foreach and if-else to work properly, added verbiage for a failed to connect error, added usage of $global:DefaultVIServer

.OUTPUTS
    HTML Report:  SnapshotReport.HTML
#>

#make sure appended report variables are cleared before run
$SnapReport = @()
$ConsReport = @()

#Initialize logging
$Logging_ScriptRootPath = Split-Path $script:MyInvocation.MyCommand.Path
$Logging_ExecutionTime = Get-Date -Format yyyy-MM-dd-HHmmss
$Logging_ExecutionDirectory = $Logging_ScriptRootPath + "\Logging\"
$Logging_ExecutionLogFile = $Logging_ExecutionDirectory + $Logging_ExecutionTime + "_vSphere_Report.log"
$SnapReportFile = $Logging_ExecutionDirectory + $Logging_ExecutionTime + "_SnapReport.html"
$ConsReportFile = $Logging_ExecutionDirectory + $Logging_ExecutionTime + "_ConsReport.html"
$vSphereReportFile = $Logging_ExecutionDirectory + $Logging_ExecutionTime + "_vSphere_Report.html"

#Create logging folder path if needed
If (!(Test-Path -Path $Logging_ExecutionDirectory)){
    Try{
        $NewDirectory = New-Item -ItemType directory -Path $Logging_ExecutionDirectory -ErrorAction Stop
    }
    Catch{
        Throw "Could not create logging directory: " + $Logging_ExecutionDirectory
    }
}

#region Functions

Function Set-AlternatingRows {
    [CmdletBinding()]
         Param(
             [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
             [object[]]$HTMLDocument,
      
             [Parameter(Mandatory=$True)]
             [string]$CSSEvenClass,
      
             [Parameter(Mandatory=$True)]
             [string]$CSSOddClass
         )
     Begin {
         $ClassName = $CSSEvenClass
     }
     Process {
         [string]$Line = $HTMLDocument
         $Line = $Line.Replace("<tr>","<tr class=""$ClassName"">")
         If ($ClassName -eq $CSSEvenClass)
         {    $ClassName = $CSSOddClass
         }
         Else
         {    $ClassName = $CSSEvenClass
         }
         $Line = $Line.Replace("<table>","<table width=""50%"">")
         Return $Line
     }
}

Function Get-SnapshotCreator {
    Param (
        [string]$VM,
        [datetime]$Created
    )

    (Get-VIEvent -Entity $VM -Types Info -Start $Created.AddSeconds(-30) -Finish $Created.AddSeconds(30) | Where-Object FullFormattedMessage -eq "Task: Create virtual machine snapshot" | Select-Object -ExpandProperty UserName).Split("\")[-1]
}

Function Script_Logging{
Param(
    [Parameter(Position=0,
        Mandatory=$True,
        ValueFromPipeLine=$True,
        ValueFromPipeLineByPropertyName=$True)]
        [String]$LogString
    )
    $CurrentTime = Get-Date -Format "HH:mm:ss.fff"
    Add-Content $Logging_ExecutionLogFile -Value ($CurrentTime + $LogString) -Encoding Unicode -ErrorAction Stop
}

#endregion

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;table-layout: auto;width: 100%}
TR:Hover TD {Background-Color: #C1D5F8;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title>
VMware vSphere Report
</title>
"@

$vcenters = @(
'vcenter1.fqdn',
'vcenter2.fqdn',
'vcenter3.fqdn',
'vcenter4.fqdn',
'vcenter5.fqdn'
)

Script_Logging "`tvCenter list: $vcenters"

#Import VMware modules
Find-Module -Name VMware.* | Import-Module -ErrorAction SilentlyContinue
$LoadedModules = Get-Module -Name VMware.*
Script_Logging "`tVMware Powershell modules loaded: $LoadedModules"

Foreach ($vcenter in $vcenters){
Script_Logging "`tAttempting connection to $vcenter"
Connect-VIServer $vcenter
If($global:defaultviserver){
Script_Logging "`tSuccessfully connected to $global:DefaultVIServer.Name"
Script_Logging "`tLooking for snapshots on $global:DefaultVIServer.Name"
$Snaps = Get-VM -Server $global:DefaultVIServer.Name | Get-Snapshot | Select-Object @{N="vCenter";E={$global:DefaultVIServer.Name}},VM,PowerState,Name,Description,@{Name="SizeGB";Expression={ [math]::Round($_.SizeGB,2) }},@{Name="Creator";Expression={ Get-SnapshotCreator -VM $_.VM -Created $_.Created }},Created

    If (!($Snaps)){
    $SnapReport += [PSCustomObject]@{vCenter="#boldgreen No snapshots found on $global:DefaultVIServer boldgreen#"
    VM=""
    PowerState=""
    Name=""
    Description=""
    SizeGB=""
    Creator=""
    Created=""}
    Script_Logging "`tNo snapshots found on $global:DefaultVIServer.Name"
        }
    Else{
    $SnapReport += $Snaps
    $SnapsString = $Snaps | Out-String
    Script_Logging "`tFound snapshots: $SnapsString"
        }
    Script_Logging "`tLooking for disks requiring consolidation on $global:DefaultVIServer.Name"
    $Cons = Get-VM -Server $global:DefaultVIServer.Name | Where-Object {$_.Extensiondata.Runtime.ConsolidationNeeded} | Select-Object Name, VMHost, @{Name = 'ConsolidationNeeded' ;Expression = {$_.Extensiondata.Runtime.ConsolidationNeeded}}
        If(!($Cons)){
        $ConsReport += [PSCustomObject]@{vCenter="#boldgreen No VM disks requiring consolidation found on $global:DefaultVIServer boldgreen#"
        Name=""
        VMHost=""
        PowerState=""
        ConsolidationNeeded=""}
        Script_Logging "`tNo disks requiring consolidation on $global:DefaultVIServer.Name"
            }
        Else {$ConsReport += $Cons
        $ConsString = $Cons | Out-String
        Script_Logging "`tFound disks requiring consolidation: $ConsString"
            }
    Disconnect-VIServer -Force -Confirm:$false

    If(!($global:defaultviserver)){
    Script_Logging "`tSuccessfully disconnected from $vcenter"}
    Else{
    Script_Logging "`tCannot get away from $vcenter!"}
        }
Else{
$SnapReport += [PSCustomObject]@{vCenter="#boldred Failed to connect to $vcenter boldred#"
VM=""
PowerState=""
Name=""
Description=""
SizeGB=""
Creator=""
Created=""}
$ConsReport += [PSCustomObject]@{vCenter="#boldred Failed to connect to $vcenter boldred#"
Name=""
VMHost=""
PowerState=""
ConsolidationNeeded=""}
Script_Logging "`tFailed to connect to $vcenter"
    }
}

#$PathToReport = "C:\Files\Scripts\WorkInProgress\"

$SnapReport = $SnapReport | ConvertTo-Html -Head $Header -PreContent "<p><h2><font color=blue>vCenter Snapshots Report</font></h2></p>" | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd
$SnapReport = $SnapReport -replace "#boldgreen","<font color=green><b>"
$SnapReport = $SnapReport -replace "boldgreen#","</b></font>"
$SnapReport = $SnapReport -replace "#boldred","<font color=red><b>"
$SnapReport = $SnapReport -replace "boldred#","</b></font>"
$SnapReport = $SnapReport -replace "<td>vcenter1.fqdn</td>","<td><font color=darkbluevcenter1.fqdn</font></td>"
$SnapReport = $SnapReport -replace "<td>vcenter2.fqdn</td>","<td><font color=darkred>vcenter2.fqdn</font></td>"
$SnapReport = $SnapReport -replace "<td>vcenter3.fqdn</td>","<td><font color=teal>vcenter3.fqdn</font></td>"
$SnapReport = $SnapReport -replace "<td>vcenter4.fqdn</td>","<td><font color=darkslategray>vcenter4.fqdn</font></td>"
$SnapReport = $SnapReport -replace "<td>vcenter5.fqdn</td>","<td><font color=slategray>vcenter5.fqdn</font></td>"
#$SnapReport | Out-File $SnapReportFile

$ConsReport = $ConsReport | ConvertTo-Html -Head $Header -PreContent "<p><h2><font color=blue>vCenter Disk Consolidation Report</font></h2></p>" | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd
$ConsReport = $ConsReport -replace "#boldgreen","<font color=green><b>"
$ConsReport = $ConsReport -replace "boldgreen#","</b></font>"
$ConsReport = $ConsReport -replace "#boldred","<font color=red><b>"
$ConsReport = $ConsReport -replace "boldred#","</b></font>"
$ConsReport = $ConsReport -replace "VMHost","VM"
$ConsReport = $ConsReport -replace "<td>vcenter1.fqdn</td>","<td><font color=darkblue>vcenter1.fqdn</font></td>"
$ConsReport = $ConsReport -replace "<td>vcenter2.fqdn</td>","<td><font color=darkred>vcenter2.fqdn</font></td>"
$ConsReport = $ConsReport -replace "<td>vcenter3.fqdn</td>","<td><font color=teal>vcenter3.fqdn</font></td>"
$ConsReport = $ConsReport -replace "<td>vcenter4.fqdn</td>","<td><font color=darkslategray>vcenter4.fqdn</font></td>"
$ConsReport = $ConsReport -replace "<td>vcenter5.fqdn</td>","<td><font color=slategray>vcenter5.fqdn</font></td>"
#$ConsReport | Out-File $ConsReportFile

$footer = "Process run from $env:computername with account $env:username."
$Report = $SnapReport + '<p><p>' + $ConsReport + '<p><p>' + $footer
$date = Get-Date -Format yyyy-MM-dd
$Report | Out-File $vSphereReportFile

$To = 'to@email'
$From = 'from@email'
$SMTPServer = 'smtp.server'

$MailSplat = @{
    To         = $To
    From       = $From
    Subject    = VMware vSphere Report ($date)"
    Body       = ($Report | Out-String)
    BodyAsHTML = $true
    SMTPServer = $SMTPServer
}

Send-MailMessage @MailSplat
Script_Logging "`tReporting completed"
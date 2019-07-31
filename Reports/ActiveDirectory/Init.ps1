<#
.SYNOPSIS  
    Init.ps1 installs all Roles and Features that are required to run ActiveDirectory reports.
.DESCRIPTION
    Init.ps1 installs all Roles and Features that are required to run ActiveDirectory reports.
.NOTES
    Version:        0.2.0
    Author:         Daniel Oldberg
    Twitter:        @danieloldberg
    Github:         danieloldberg
    Credits:        Iain Brighton (@iainbrighton) - PScribo module
                    Carl Webster (@carlwebster) - Documentation Script Concept
#>

$InstalledModules = Get-Module -ListAvailable

If($InstalledModules.Name -notcontains  "pscribo"){
    Try{
        Install-Module -Name "Pscribo" -Scope CurrentUser -Force -Confirm:$false -ErrorAction Stop
    }
    Catch{
        Write-Error "An error occured while trying to install Pscribo."
    }
}

# Get Operating System Product Name
$OSVersion = (get-itemproperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductName).ProductName

# Check if the machine is a Windows Server
If($OSVersion -like "*Server*"){
    If($InstalledModules.Name -notcontains "dfsn"){
        Try{
            Install-WindowsFeature -Name "RSAT-DFS-Mgmt-Con" -ErrorAction Stop
        }
        Catch{
            Write-Error "An error occured while trying to install Dfsn Management Tools."
        }
    }

    If($InstalledModules.Name -notcontains "ActiveDirectory"){
        Try{
            Install-WindowsFeature -Name "RSAT-ADDS" -IncludeAllSubFeature -IncludeManagementTools -ErrorAction Stop
        }
        Catch{
            Write-Error "An error occured while trying to install Active Directory Management Tools."
        }
    }

    If($InstalledModules.Name -notcontains "GroupPolicy"){
        Try{
            Install-WindowsFeature -Name "GPMC" -IncludeAllSubFeature -IncludeManagementTools -ErrorAction Stop
        }
        Catch{
            Write-Error "An error occured while trying to install Group Policy Management Tools."
        }
    }

    If($InstalledModules.Name -notcontains "DnsServer"){
        Try{
            Install-WindowsFeature -Name "RSAT-DNS-Server" -IncludeAllSubFeature -IncludeManagementTools -ErrorAction Stop
        }
        Catch{
            Write-Error "An error occured while trying to install Dns Server Management Tools."
        }
    }
}
# Assume is a desktop
Else{
    If($InstalledModules.Name -notcontains  "dfsn"){
        Write-Warning ("Unable to install Dfsn RSAT tools. Please visit https://www.google.com/search?q=" + $OSVersion.Replace(" ","+") + "+rsat+tools and install manually")
    }

    If($InstalledModules.Name -notcontains  "ActiveDirectory"){
        Write-Warning ("Unable to install Active Directory RSAT tools. Please visit https://www.google.com/search?q=" + $OSVersion.Replace(" ","+") + "+rsat+tools and install manually")
    }

    If($InstalledModules.Name -notcontains  "GroupPolicy"){
        Write-Warning ("Unable to install Group Policy RSAT tools. Please visit https://www.google.com/search?q=" + $OSVersion.Replace(" ","+") + "+rsat+tools and install manually")
    }

    If($InstalledModules.Name -notcontains  "DnsServer"){
        Write-Warning ("Unable to install Dns Server RSAT tools. Please visit https://www.google.com/search?q=" + $OSVersion.Replace(" ","+") + "+rsat+tools and install manually")
    }
}
#requires -Modules @{ModuleName="PScribo";ModuleVersion="0.7.24"},ActiveDirectory,GroupPolicy

<#
.SYNOPSIS  
    PowerShell script to document the configuration of Microsoft Active Directory in Word/HTML/XML/Text formats
.DESCRIPTION
    Documents the configuration of Microsoft Active Directory in Word/HTML/XML/Text formats using PScribo.
.NOTES
    Version:        0.1.0
    Author:         Daniel Oldberg
    Twitter:        @danieloldberg
    Github:         danieloldberg
    Credits:        Iain Brighton (@iainbrighton) - PScribo module
                    Jake Rutski (@jrutski) - VMware vSphere Documentation Script Concept
.LINK
    https://github.com/tpcarman/As-Built-Report
    https://github.com/iainbrighton/PScribo
#>

#region Configuration Settings
#---------------------------------------------------------------------------------------------#
#                                    CONFIG SETTINGS                                          #
#---------------------------------------------------------------------------------------------#
# Clear variables
$DomainControllers = @()

# If custom style not set, use VMware style
if (!$StyleName) {
    & "$PSScriptRoot\..\..\Styles\VMware.ps1"
}

#endregion Configuration Settings

#region Script Functions
#---------------------------------------------------------------------------------------------#
#                                    SCRIPT FUNCTIONS                                         #
#---------------------------------------------------------------------------------------------#


#endregion Script Functions

#region Script Body
#---------------------------------------------------------------------------------------------#
#                                         SCRIPT BODY                                         #
#---------------------------------------------------------------------------------------------#

foreach ($Forest in $Target) {
    
    Section -Style Heading1 "Forest Summary" {
        
        if ($InfoLevel.Forest -ge 1) {

        $ForestObject = Get-ADForest -Identity $Forest -Credential $Credentials

            Paragraph "Active Directory has a forest name $Forest. Following table contains forest summary with important information:"
            
            Section -Style Heading2 "FSMO Servers" {

                Paragraph "Following table contains FSMO servers"

                $ForestObject | Select-Object SchemaMaster,DomainNamingMaster | Table -Name "FSMO Roles" -List

            }

            Section -Style Heading2 "Optional Forest Features" {

                Paragraph "Following table contains optional forest features"

                

            }

            Section -Style Heading2 "UPN Suffixes" {

                Paragraph "Following UPN suffixes were created in this forest:"

                

            }



            

        }
    }

    Section -Style Heading1 "Forest Sites" {

    }

    Section -Style Heading1 "Forest Subnets" {

    }

    Section -Style Heading1 "Forest Site Links" {

    }

    # Loop all domains in forest 
    ForEach($Domain in ($ForestObject.Domains)){

        Section -Style Heading1 "Domain - $Domain" {

            # Collect global domain information
            $DomainObject = Get-ADDomain -Identity $Domain -Credential $Credentials

            Section -Style Heading2 "Domain - $Domain - Domain Controllers" {



            }

            Section -Style Heading2 "Domain - $Domain - FSMO Roles" {
                
                $DomainObject | Select-Object InfrastructureMaster,PDCEmulator,RIDMaster | Table -Name "$Domain FSMO Roles" -List

            }

            Section -Style Heading2 "Domain - $Domain - Password Policies" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - Fine Grained Password Policies" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - Group Policies" {
                
                Try{
                    $DomainGPOs = Get-GPO -domain $Domain -All -ErrorAction Stop

                }
                Catch{
                    Write-Verbose "Unable to collect GPO information for $domain. This is probably due to missing permissions or client machine in another domain"
                    Paragraph "Unable to collect GPO information for $domain. This is probably due to missing permissions or client machine in another domain" -Color Red
                    Continue
                }

                
            }
                

            }

            Section -Style Heading2 "Domain - $Domain - Group Policies Details" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - Group Policies ACL" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - DNS A/SRV Records" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - Trusts" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - Organizational Units" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - Domain Administrators" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - Enterprise Administrators" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - Users Count" {
                
                

            }

            Section -Style Heading2 "Domain - $Domain - GPP Drive Maps" {
                
                # If we were able to retrieve domain GPO objects
                If($DomainGPOs){
                    
                    # Thanks to Johan Dahlbom @ https://365lab.net/2013/12/31/getting-all-gpp-drive-maps-in-a-domain-with-powershell/
                    foreach ($Policy in $DomainGPOs){
            
                        $GPOID = $Policy.Id
                        $GPODom = $Policy.DomainName
                        $GPODisp = $Policy.DisplayName
        
                        if (Test-Path "\\$($GPODom)\SYSVOL\$($GPODom)\Policies\{$($GPOID)}\User\Preferences\Drives\Drives.xml")
                        {
                            [xml]$DriveXML = Get-Content "\\$($GPODom)\SYSVOL\$($GPODom)\Policies\{$($GPOID)}\User\Preferences\Drives\Drives.xml"
        
                            foreach ( $drivemap in $DriveXML.Drives.Drive )

                            {
                                New-Object PSObject -Property @{
                                GPOName = $GPODisp
                                DriveLetter = $drivemap.Properties.Letter + ":"
                                DrivePath = $drivemap.Properties.Path
                                DriveAction = $drivemap.Properties.action.Replace("U","Update").Replace("C","Create").Replace("D","Delete").Replace("R","Replace")
                                DriveLabel = $drivemap.Properties.label
                                DrivePersistent = $drivemap.Properties.persistent.Replace("0","False").Replace("1","True")
                                DriveFilterGroup = $drivemap.Filters.FilterGroup.Name
                                }
                            }
                        }

                    }

                }
                # If Domain GPOs were unable to be retrieved.
                Else{

                    Write-Verbose "Unable to collect GPP Drive Maps for $domain. This is probably due to not being able to retrieve GPO objects"
                    Paragraph "Unable to collect GPP Drive Maps for $domain. This is probably due to not being able to retrieve GPO objects" -Color Red

                }

            }






        
        }
    
}


#endregion Script Body

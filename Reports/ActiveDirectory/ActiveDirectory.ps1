#requires -Modules @{ModuleName="PScribo";ModuleVersion="0.7.24"},ActiveDirectory,GroupPolicy,dfsn

<#
.SYNOPSIS  
    PowerShell script to document the configuration of Microsoft Active Directory in Word/HTML/XML/Text formats
.DESCRIPTION
    Documents the configuration of Microsoft Active Directory in Word/HTML/XML/Text formats using PScribo.

    The script is meant to be run on a management machine belonging to the same domain.
    This is because of the Get-GPO cmdlet that is not possible to be used in a remoting purpose.
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

Import-Module ActiveDirectory
Import-Module GroupPolicy
Import-Module dfsn

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

    $ForestObject = Get-ADForest -Identity $Forest -Credential $Credentials
    
    Section -Style Heading1 "Forest Summary" {
        
        if ($InfoLevel.Forest -ge 1) {

            Paragraph "Active Directory has a forest name $Forest. Following table contains forest summary with important information:"
            
            Section -Style Heading2 "FSMO Servers" {
                Paragraph "Following table contains FSMO servers"
                $ForestObject | Select-Object SchemaMaster,DomainNamingMaster | Table -Name "FSMO Servers" -List
            }

            Section -Style Heading2 "Optional Forest Features" {
                Paragraph "Following table contains optional forest features"
                $ForestOptionalFeatureObject = Get-ADOptionalFeature -Filter * -Server $Forest -Credential $Credentials
                # Check if Recycle bin is enabled
                If(($ForestOptionalFeatureObject | Where{$_.Name -like "*Recycle Bin Feature*"}).EnabledScopes){
                    $RecycleBinStatus = $True
                }
                Else {
                    $RecycleBinStatus = $False
                }
                # Check if Privileged Access Management Feature is enabled
                If(($ForestOptionalFeatureObject | Where{$_.Name -like "*Privileged Access Management Feature*"}).EnabledScopes){
                    $PAMStatus = $True
                }
                Else {
                    $PAMStatus = $False
                }
                $ForestOptionalFeatureHash = [Ordered]@{
                    "Recycle Bin Enabled"                           = $RecycleBinStatus
                    "Privileged Access Management Feature Enabled"  = $PAMStatus
                }
                New-Object PSObject -Property $ForestOptionalFeatureHash | Table -Name "Forest UPN suffixes" -List -ColumnWidths 75,25
            }

            Section -Style Heading2 "Forest UPN Suffixes" {
                Paragraph "Following UPN suffixes were created in this forest:"
                $UPNSuffixObject = @()
                ForEach($UPNSuffix in $ForestObject.UPNSuffixes){
                    $UPNSuffixObject += New-Object PSObject -Property @{
                        UPNSuffix = $UPNSuffix
                    }
                }
                $UPNSuffixObject | Table -Name "Forest UPN suffixes"

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

            Section -Style Heading2 "Domain Controllers" {
                $DomainDCs = Get-ADGroupMember 'Domain Controllers' -Credential $Credentials -Server $Domain | Get-ADDomainController
                $DomainDCs | Select-Object -Property HostName,@{Name="Read Only DC";Expression={$_."IsReadOnly"}},@{Name="Global Catalog";Expression={$_."IsGlobalCatalog"}},IPv4Address,OperatingSystem,Site | Table -Name "$Domain Domain Controllers"
            }

            Section -Style Heading2 "UPN Suffixes" {
                Paragraph "Following UPN suffixes were created in this domain:"
                $DomainUPNDN = ("cn=Partitions,cn=Configuration," + $DomainObject.DistinguishedName)
                $UPNSuffixes = Get-ADObject -Identity $DomainUPNDN -Properties upnsuffixes -Credential $Credentials -Server $Domain | Select-Object -ExpandProperty upnsuffixes
                $UPNSuffixObject = @()
                ForEach($UPNSuffix in $UPNSuffixes){
                    $UPNSuffixObject += New-Object PSObject -Property @{
                        UPNSuffix = $UPNSuffix
                    }
                }
                $UPNSuffixObject | Table -Name "Domain UPN suffixes"
            }

            Section -Style Heading2 "FSMO Servers" {
                
                $DomainObject | Select-Object InfrastructureMaster,PDCEmulator,RIDMaster | Table -Name "$Domain FSMO Servers" -List

            }

            Section -Style Heading2 "Password Policies" {
                
                # Get-ADDefaultDomainPasswordPolicy

            }

            Section -Style Heading2 "Fine Grained Password Policies" {
                
                # Get-ADFineGrainedPasswordPolicy

            }

            Section -Style Heading2 "Group Policies" {
                
                Try{
                    $DomainGPOs = Get-GPO -domain $Domain -All -ErrorAction Stop

                    $DomainGPOs | Select-Object DisplayName,GpoStatus,Description,CreationTime |Table -Name "$Domain Group Policies"

                }
                Catch{
                    Write-Verbose "Unable to collect GPO information for $domain. This is probably due to missing permissions or client machine in another domain"
                    Paragraph "Unable to collect GPO information for $domain. This is probably due to missing permissions or client machine in another domain" -Color Red
                    Return
                }
                
                
            }
                

            Section -Style Heading2 "Group Policies Details" {
            
            }

            Section -Style Heading2 "Group Policies ACL" {
            
            }

            Section -Style Heading2 "DNS A/SRV Records" {
        
            }

            Section -Style Heading2 "Trusts" {
                
            }

            Section -Style Heading2 "Organizational Units" {
                Paragraph "Following table contains all OU's created in $Domain"
                $DomainOUs = Get-ADOrganizationalUnit -Server $Domain -Credential $Credentials -Properties * -filter *
                $DomainOUs | Select-Object CanonicalName,ManagedBy,@{Name="Protected";Expression={$_."ProtectedFromAccidentalDeletion"}},Created | Table -Name "$Domain Organizational Units"
            }

            Section -Style Heading2 "Domain Administrators" {
                Paragraph "Following users have highest priviliges and are able to control a lot of Windows resources."
                $EnterpriseAdmins = Get-ADGroupMember 'Domain Admins' -Credential $Credentials -Server $Domain | Get-ADUser
                $EnterpriseAdmins | Select-Object Enabled, Name, SamAccountName, UserPrincipalName | Table -Name "$Domain Domain Admins" -ColumnWidths 15,20,30,35
            }

            Section -Style Heading2 "Enterprise Administrators" {
                Paragraph "Following users have highest priviliges across Forest and are able to control a lot of Windows resources."
                $EnterpriseAdmins = Get-ADGroupMember 'Enterprise Admins' -Credential $Credentials -Server $Domain | Get-ADUser
                $EnterpriseAdmins | Select-Object Enabled, Name, SamAccountName, UserPrincipalName | Table -Name "$Domain Enterprise Admins" -ColumnWidths 15,20,30,35
            }

            Section -Style Heading2 "Users Count" {
                $UserObject = Get-ADuser -Credential $Credentials -Filter * -Server $Domain
                
                $UserHash = [Ordered]@{
                    "Users Count Incl. System"              = $UserObject.Count
                    "Users Count"                           = $PAMStatus
                    "Users Expired"                         = $(Search-ADAccount -AccountExpired -Credential $Credentials -Server $Domain).Count
                    "Users Expired Incl. Disabled"          = $PAMStatus
                    "Users Never Expiring"                  = $($UserObject | Where{$_.PasswordNeverExpires -eq $True}).Count
                    "Users Never Expiring Incl. Disabled"   = $PAMStatus
                    "Users System Accounts"                 = $PAMStatus

                }
                <#
                Users Count Incl. System	36
                Users Count	33
                Users Expired	1
                Users Expired Incl. Disabled	3
                Users Never Expiring	22
                Users Never Expiring Incl. Disabled	22
                Users System Accounts	3
                #>

            }

            Section -Style Heading2 "GPP Drive Maps" {
                
                # If we were able to retrieve domain GPO objects
                If($DomainGPOs -ne $Null){
                    
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
            
            Section -Style Heading2 "DFS Namespaces" {
                
                # Check to se if DFS module is working
                Try{

                    $DFSnRoots = Get-DfsnRoot -domain $Domain

                    ForEach($DFSnRoot in $DFSnRoots){
    
                        
                        If($DFSRoot.Flags -like "*AccessBased Enumeration*"){
                            $DFSRoot = Add-Member -InputObject $DFSRoot -MemberType NoteProperty -Name "AccessBased Enumeration" -Value "True" -PassThru
                        }
                    
                        Section -Style Heading3 $DFSnRoot.Path
                    
                        $DFSnFolders = Get-DfsnFolder -Path ($($DFSnRoot.Path) + "\*")
                        
                        ForEach($DFSnFolder in $DFSnFolders){
                    
                            $DFSnFolder | Table -Name "$DFSnRoot DFSn Folder"
                    
                            $DFSnFolderTargets = Get-DfsnFolderTarget -Path $DFSnFolder.Path
                    
                            ForEach($DFSnFolderTarget in $DFSnFolderTargets){
                    
                                $DFSnFolderTarget | Table -Name "$DFSnFolder DFSn Folder Targets"
                    
                            }
                    
                        }
                        
                    }

                }
                Catch{

                    Write-Verbose "Unable to collect DFS information for $domain. This is probably due to not DFSn module not installed or client not being inside the same domain."
                    Paragraph "Unable to collect DFS information for $domain. This is probably due to not DFSn module not installed or client not being inside the same domain." -Color Red
                    Return

                }
             
            }

        }
        
    }
    
}


#endregion Script Body

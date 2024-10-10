<#
.SYNOPSIS
    This file is part of Identity and Access Services - Privileged Access.

    This script backs up and restores Active Directory objects.

.DESCRIPTION
    This script backs up and restores Active Directory objects.
    Requirements to run the script are the following:
    - Windows 6.3+,
    - Windows PowerShell 4+,
    - user account with permissions to perform operations in Active Directory.

.PARAMETER Backup
    Enables the Backup mode.

.PARAMETER Restore
    Enables the Restore mode.

.PARAMETER SettingsXml
    Path to the Settings XML file.

.PARAMETER OutputFolder
    Path to a folder where items will be backed up. If not specified, a folder
    named Managed-AdObjects will be created into the user's temporary folder
    defined by the environment variable TEMP.

    Note: full path must not exceed 248 characters. If GPO parameter is used,
    the full path must not exceed 150 characters.

.PARAMETER OU
    In Backup mode, backs up Organizational Unit objects.

    In Restore mode, restores Organizational Unit objects.

.PARAMETER User
    In Backup mode, backs up User objects.

    In Restore mode, restores User objects.

.PARAMETER Computer
    In Backup mode, backs up Computer objects.

    In Restore mode, restores Computer objects.

.PARAMETER Group
    In Backup mode, backs up Group objects.

    In Restore mode, restores Group objects.

.PARAMETER WmiFilter
    In Backup mode, backs up WMI Filter objects.

    In Restore mode, restores WMI Filter objects.

.PARAMETER Admx
    In Backup mode, backs up ADMX and ADML files from the Central Store.

    In Restore mode, restores ADMX and ADML files to the Central Store.

.PARAMETER GPO
    In Backup mode, backs up Group Policy objects regarding the value of this
    parameter:
    - All: all Group Policy Objects are backed up.
    - LinkedOnly: only Group Policy Objects linked to an Organizational Unit
      are backed up.
    - None: Group Policy Objects are not backed up. This is the same as not
      using this parameter.

    Note: Default Domain Policy and Default Domain Controller Policy are not
    backed up by the script.

    In Restore mode, restores Group Policy Objects regarding the value of this
    parameter:
    - All: all Group Policy Objects are restored.
    - LinkedOnly: only Group Policy Objects that ware linked to an
      Organizational Unit are restored.
    - None: Group Policy Objects are not restored. This is the same as not
      using this parameter and it is the default value.

.PARAMETER GpoLinks
    Restores links of Group Policy Objects regarding the value of this
    parameter:
    - DontLink: restored Group Policy Objects are not linked. This is the
      default value.
    - LinkDisabled: restored Group Policy Objects are linked to an
      Organizational Unit object, or the domain root, but the link is disabled.
    - LinkEnabled: Restored Group Policy Objects are linked to an
      Organizational Unit object, or the domain root, and the link is enabled.

.PARAMETER GpoReports
    If true, HTML reports are generated for each backed-up GPOs. Reports are
    located into the Backup folder under Reports folder.

.PARAMETER Permissions
    In Backup mode, backs up permissions of all objects specified by the
    script's parameters.

    In Restore mode, restores permissions of all objects specified by the
    script's parameters.

.PARAMETER RedirectContainers
    This parameter only works in combination with OU parameter.

    In Backup mode, backs up redirection of system containers.
    In Restore mode, restores redirection of system containers.

.PARAMETER Scope
    Reduces the scope of the script to a specific container specified by the
    distinguishedName. By default, the scope is defined to the root of the
    Active Directory domain.

    Note: Active Directory objects not in an Organizational Unit are ignored by
    the script. For example, objects in the Users container are not backed up.

.PARAMETER Server
    Active Directory Domain Controller to target for all operations. By
    default, the Primary Domain Controller is selected.

.PARAMETER Credential
    Alternate credential to use for all operations in Active Directory. By
    default, the current user credential is used.

.PARAMETER Force
    If an item already exists, it will be overwritten. By default, this
    parameter is not enabled: existing items are not updated or overwritten.

.PARAMETER Confirm
    This parameter enabled by default pause the script execution and wait for a
    confirmation to continue.

.PARAMETER LogFile
    Path to the log file to create.

.PARAMETER LogFormat
    Format of the log file generated. If the parameter LogFile is not used,
    this parameter is ignored.

    Supported values are: CMTrace, Csv, Html, Json, Xml.

.OUTPUTS
    True on success, false otherwise.

.EXAMPLE
    PS> .\Manage-AdObjects.ps1 -Backup -OutputFolder 'C:\path\Backup' `
                               -OU -User -Computer -Group -WmiFilter -GPO All `
                               -Permissions -Server 'dc.contoso.com'

    Backs up all AD objects and their permissions targeting a specific domain
    controller. All files are stored in the specified folder.

.EXAMPLE
    PS> $cred = Get-Credential
    PS> .\Manage-AdObjects.ps1 -Restore -SettingsXml 'C:\path\settings.xml' `
                               -OU -User -Computer -Group -WmiFilter -GPO All `
                               -Permissions -Credential $cred `
                               -Server 'dc.contoso.com'

    Restores all AD objects and permissions from an Xml file using alternate
    credentials and targeting a specific domain controller.

.NOTES
    Authors
        [GS] Gregory Schiro <gregory.schiro@microsoft.com>

    2021-09-20, version 1.0.0
        [GS] First release.

    The sample scripts provided here are not supported under any Microsoft
    standard support program or service. All scripts are provided AS IS without
    warranty of any kind. Microsoft further disclaims all implied warranties
    including, without limitation, any implied warranties of merchantability or
    of fitness for a particular purpose. The entire risk arising out of the use
    or performance of the sample scripts and documentation remains with you. In
    no event shall Microsoft, its authors, or anyone else involved in the
    creation, production, or delivery of the scripts be liable for any damages
    whatsoever (including, without limitation, damages for loss of business
    profits, business interruption, loss of business information, or other
    pecuniary loss) arising out of the use of or inability to use the sample
    scripts or documentation, even if Microsoft has been advised of the
    possibility of such damages.
#>

#Requires -Version 4.0

#region Parameters

[CmdletBinding(DefaultParameterSetName='Backup')]
param (
    [Parameter(Mandatory=$true, ParameterSetName='Backup')]
    [Switch]$Backup = $false,

    [Parameter(Mandatory=$true, ParameterSetName='Restore')]
    [Switch]$Restore = $false,

    [Parameter(Mandatory=$true, ParameterSetName='Restore')]
    [String]$SettingsXml = $null,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [String]$OutputFolder = $null,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [Switch]$OU = $false,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [Switch]$User = $false,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [Switch]$Computer = $false,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [Switch]$Group = $false,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [Switch]$WmiFilter = $false,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [Switch]$Admx = $false,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [ValidateSet('All', 'LinkedOnly', 'None')]
    [String]$GPO = 'None',

    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [ValidateSet('DontLink', 'LinkDisabled', 'LinkEnabled')]
    [String]$GpoLinks = 'DontLink',

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Switch]$GpoReports = $null,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [Switch]$Permissions = $false,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [Parameter(Mandatory=$false, ParameterSetName='Restore')]
    [Switch]$RedirectContainers = $null,

    [Parameter(Mandatory=$false, ParameterSetName='Backup')]
    [String]$Scope = $null,

    [Parameter(Mandatory=$false)]
    [String]$Server = $null,

    [Parameter(Mandatory=$false)]
    [Management.Automation.PSCredential]$Credential = $null,

    [Parameter(Mandatory=$false)]
    [Switch]$Force = $false,

    [Parameter(Mandatory=$false)]
    [Switch]$Confirm = $true,

    [Parameter(Mandatory=$false)]
    [String]$LogFile = $null,

    [Parameter(Mandatory=$false)]
    [ValidateSet('CMTrace', 'Csv', 'Html', 'Json', 'Xml')]
    [String]$LogFormat = 'CMTrace'
)

#endregion

Set-StrictMode -Version Latest

#region Variables

# Current version of the script.
Set-Variable -Scope Script -Option Constant -Name ScriptVersion -Value '1.0.0'

# Xml nodes' names used in Settings Xml file (values are case sensitive).
Set-Variable -Scope Script -Option Constant -Name DefAccessControl -Value 'accessControl'
Set-Variable -Scope Script -Option Constant -Name DefAccountExpirationDate -Value 'accountExpirationDate'
Set-Variable -Scope Script -Option Constant -Name DefAccountNotDelegated -Value 'accountNotDelegated'
Set-Variable -Scope Script -Option Constant -Name DefAccountPassword -Value 'accountPassword'
Set-Variable -Scope Script -Option Constant -Name DefRights -Value 'rights'
Set-Variable -Scope Script -Option Constant -Name DefAllowRevPwdEncryption -Value 'allowReversiblePasswordEncryption'
Set-Variable -Scope Script -Option Constant -Name DefCannotChangePassword -Value 'cannotChangePassword'
Set-Variable -Scope Script -Option Constant -Name DefCategory -Value 'category'
Set-Variable -Scope Script -Option Constant -Name DefChangePasswordAtLogon -Value 'changePasswordAtLogon'
Set-Variable -Scope Script -Option Constant -Name DefComputer -Value 'computer'
Set-Variable -Scope Script -Option Constant -Name DefComputers -Value 'computers'
Set-Variable -Scope Script -Option Constant -Name DefConfiguration -Value 'configuration'
Set-Variable -Scope Script -Option Constant -Name DefContainer -Value 'container'
Set-Variable -Scope Script -Option Constant -Name DefDescription -Value 'description'
Set-Variable -Scope Script -Option Constant -Name DefDisplayName -Value 'displayName'
Set-Variable -Scope Script -Option Constant -Name DefDistinguishedName -Value 'distinguishedName'
Set-Variable -Scope Script -Option Constant -Name DefEnabled -Value 'enabled'
Set-Variable -Scope Script -Option Constant -Name DefEnforced -Value 'enforced'
Set-Variable -Scope Script -Option Constant -Name DefGivenName -Value 'givenName'
Set-Variable -Scope Script -Option Constant -Name DefGpoBackupFolder -Value 'gpoBackupFolder'
Set-Variable -Scope Script -Option Constant -Name DefGpoInheritanceBlocked -Value 'gpoInheritanceBlocked'
Set-Variable -Scope Script -Option Constant -Name DefGpoStatus -Value 'gpoStatus'
Set-Variable -Scope Script -Option Constant -Name DefGroup -Value 'group'
Set-Variable -Scope Script -Option Constant -Name DefGroupPolicies -Value 'groupPolicies'
Set-Variable -Scope Script -Option Constant -Name DefGroupPolicy -Value 'groupPolicy'
Set-Variable -Scope Script -Option Constant -Name DefGroups -Value 'groups'
Set-Variable -Scope Script -Option Constant -Name DefIdentityReference -Value 'identityReference'
Set-Variable -Scope Script -Option Constant -Name DefInheritance -Value 'inheritance'
Set-Variable -Scope Script -Option Constant -Name DefInheritedObject -Value 'inheritedObject'
Set-Variable -Scope Script -Option Constant -Name DefInitials -Value 'initials'
Set-Variable -Scope Script -Option Constant -Name DefIsWellKnown -Value 'isWellKnown'
Set-Variable -Scope Script -Option Constant -Name DefKerberosEncryptionType -Value 'kerberosEncryptionType'
Set-Variable -Scope Script -Option Constant -Name DefLink -Value 'link'
Set-Variable -Scope Script -Option Constant -Name DefLinks -Value 'links'
Set-Variable -Scope Script -Option Constant -Name DefManagedBy -Value 'managedBy'
Set-Variable -Scope Script -Option Constant -Name DefMember -Value 'member'
Set-Variable -Scope Script -Option Constant -Name DefMembers -Value 'members'
Set-Variable -Scope Script -Option Constant -Name DefName -Value 'name'
Set-Variable -Scope Script -Option Constant -Name DefObject -Value 'object'
Set-Variable -Scope Script -Option Constant -Name DefOrder -Value 'order'
Set-Variable -Scope Script -Option Constant -Name DefOu -Value 'organizationalUnit'
Set-Variable -Scope Script -Option Constant -Name DefOus -Value 'organizationalUnits'
Set-Variable -Scope Script -Option Constant -Name DefParameters -Value 'parameters'
Set-Variable -Scope Script -Option Constant -Name DefPasswordNeverExpires -Value 'passwordNeverExpires'
Set-Variable -Scope Script -Option Constant -Name DefPasswordNotRequired -Value 'passwordNotRequired'
Set-Variable -Scope Script -Option Constant -Name DefPermission -Value 'permission'
Set-Variable -Scope Script -Option Constant -Name DefPermissions -Value 'permissions'
Set-Variable -Scope Script -Option Constant -Name DefProtectedFromDel -Value 'protectedFromAccidentalDeletion'
Set-Variable -Scope Script -Option Constant -Name DefSamAccountName -Value 'samAccountName'
Set-Variable -Scope Script -Option Constant -Name DefScope -Value 'scope'
Set-Variable -Scope Script -Option Constant -Name DefShowAdvView -Value 'showInAdvancedView'
Set-Variable -Scope Script -Option Constant -Name DefSmartcardLogonRequired -Value 'smartcardLogonRequired'
Set-Variable -Scope Script -Option Constant -Name DefSurname -Value 'surname'
Set-Variable -Scope Script -Option Constant -Name DefSystemContainer -Value 'systemContainer'
Set-Variable -Scope Script -Option Constant -Name DefTargetName -Value 'targetName'
Set-Variable -Scope Script -Option Constant -Name DefTrustedForDelegation -Value 'trustedForDelegation'
Set-Variable -Scope Script -Option Constant -Name DefUser -Value 'user'
Set-Variable -Scope Script -Option Constant -Name DefUserPrincipalName -Value 'userPrincipalName'
Set-Variable -Scope Script -Option Constant -Name DefUsers -Value 'users'
Set-Variable -Scope Script -Option Constant -Name DefVersion -Value 'version'
Set-Variable -Scope Script -Option Constant -Name DefWmiAuthor -Value 'wmiAuthor'
Set-Variable -Scope Script -Option Constant -Name DefWmiFilter -Value 'wmiFilter'
Set-Variable -Scope Script -Option Constant -Name DefWmiFilters -Value 'wmiFilters'
Set-Variable -Scope Script -Option Constant -Name DefWmiName -Value 'wmiName'
Set-Variable -Scope Script -Option Constant -Name DefWmiParm1 -Value 'wmiParam1'
Set-Variable -Scope Script -Option Constant -Name DefWmiParm2 -Value 'wmiParam2'

# Schema used by Settings Xml file.
Set-Variable -Scope Script -Option Constant -Name XmlSchema -Value @"
<?xml version="1.0" encoding="utf-8"?>
<xs:schema attributeFormDefault="unqualified"
           elementFormDefault="qualified"
           xmlns:xs="http://www.w3.org/2001/XMLSchema">
    <xs:simpleType name="guid">
        <xs:restriction base="xs:string">
            <xs:pattern value="[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}"/>
        </xs:restriction>
    </xs:simpleType>
    <xs:element name="$($Script:DefConfiguration)" type="configurationType"/>
    <xs:complexType name="parametersType">
        <xs:all>
            <xs:element name="$($Script:DefVersion)" minOccurs="0">
                <xs:simpleType>
                    <xs:restriction base="xs:string">
                        <xs:pattern value="(0|[1-9][0-9]*)(\.(0|[1-9][0-9]*)){0,3}"/>
                    </xs:restriction>
                </xs:simpleType>
            </xs:element>
            <xs:element type="xs:string" name="$($Script:DefGpoBackupFolder)" minOccurs="0"/>
        </xs:all>
    </xs:complexType>
    <xs:complexType name="wellKnownType">
        <xs:simpleContent>
            <xs:extension base="xs:string">
                <xs:attribute type="xs:boolean" name="$($Script:DefIsWellKnown)" use="required"/>
            </xs:extension>
        </xs:simpleContent>
    </xs:complexType>
    <xs:complexType name="permissionType">
        <xs:all>
            <xs:element name="$($Script:DefRights)">
                <xs:simpleType>
                    <xs:list>
                        <xs:simpleType>
                            <xs:restriction base="xs:string">
                                <xs:enumeration value="AccessSystemSecurity"/>
                                <xs:enumeration value="CreateChild"/>
                                <xs:enumeration value="Delete"/>
                                <xs:enumeration value="DeleteChild"/>
                                <xs:enumeration value="DeleteTree"/>
                                <xs:enumeration value="ExtendedRight"/>
                                <xs:enumeration value="GenericAll"/>
                                <xs:enumeration value="GenericExecute"/>
                                <xs:enumeration value="GenericRead"/>
                                <xs:enumeration value="GenericWrite"/>
                                <xs:enumeration value="ListChildren"/>
                                <xs:enumeration value="ListObject"/>
                                <xs:enumeration value="ReadControl"/>
                                <xs:enumeration value="ReadProperty"/>
                                <xs:enumeration value="Self"/>
                                <xs:enumeration value="Synchronize"/>
                                <xs:enumeration value="WriteDacl"/>
                                <xs:enumeration value="WriteOwner"/>
                                <xs:enumeration value="WriteProperty"/>
                            </xs:restriction>
                        </xs:simpleType>
                    </xs:list>
                </xs:simpleType>
            </xs:element>
            <xs:element name="$($Script:DefInheritance)" minOccurs="0">
                <xs:simpleType>
                    <xs:restriction base="xs:string">
                        <xs:enumeration value="All"/>
                        <xs:enumeration value="Children"/>
                        <xs:enumeration value="Descendents"/>
                        <xs:enumeration value="None"/>
                        <xs:enumeration value="SelfAndChildren"/>
                    </xs:restriction>
                </xs:simpleType>
            </xs:element>
            <xs:element type="xs:string" name="$($Script:DefObject)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefInheritedObject)" minOccurs="0"/>
            <xs:element name="$($Script:DefAccessControl)">
                <xs:simpleType>
                    <xs:restriction base="xs:string">
                        <xs:enumeration value="Allow"/>
                        <xs:enumeration value="Deny"/>
                    </xs:restriction>
                </xs:simpleType>
            </xs:element>
            <xs:element type="wellKnownType" name="$($Script:DefIdentityReference)"/>
        </xs:all>
    </xs:complexType>
    <xs:complexType name="permissionsType">
        <xs:sequence>
            <xs:element type="permissionType" name="$($Script:DefPermission)" minOccurs="0" maxOccurs="unbounded"/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name="organizationalUnitType">
        <xs:all>
            <xs:element type="xs:string" name="$($Script:DefDescription)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefDistinguishedName)"/>
            <xs:element type="xs:boolean" name="$($Script:DefGpoInheritanceBlocked)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefName)"/>
            <xs:element type="xs:boolean" name="$($Script:DefProtectedFromDel)" minOccurs="0"/>
            <xs:element name="$($Script:DefSystemContainer)" minOccurs="0">
                <xs:simpleType>
                    <xs:list>
                        <xs:simpleType>
                            <xs:restriction base="xs:string">
                                <xs:enumeration value="NTDSQuotas"/>
                                <xs:enumeration value="Microsoft"/>
                                <xs:enumeration value="ProgramData"/>
                                <xs:enumeration value="ForeignSecurityPrincipals"/>
                                <xs:enumeration value="DeletedObjects"/>
                                <xs:enumeration value="Infrastructure"/>
                                <xs:enumeration value="LostAndFound"/>
                                <xs:enumeration value="System"/>
                                <xs:enumeration value="DomainControllers"/>
                                <xs:enumeration value="Computers"/>
                                <xs:enumeration value="Users"/>
                                <xs:enumeration value="ManagedServiceAccounts"/>
                            </xs:restriction>
                        </xs:simpleType>
                    </xs:list>
                </xs:simpleType>
            </xs:element>
            <xs:element type="permissionsType" name="$($Script:DefPermissions)" minOccurs="0"/>
        </xs:all>
    </xs:complexType>
    <xs:complexType name="organizationalUnitsType">
        <xs:sequence>
            <xs:element type="organizationalUnitType" name="$($Script:DefOu)" minOccurs="0" maxOccurs="unbounded"/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name="userType">
        <xs:all>
            <xs:element type="xs:string" name="$($Script:DefAccountExpirationDate)" minOccurs="0"/>
            <xs:element type="xs:boolean" name="$($Script:DefAccountNotDelegated)" minOccurs="0"/>
            <xs:element type="xs:boolean" name="$($Script:DefAllowRevPwdEncryption)" minOccurs="0"/>
            <xs:element type="xs:boolean" name="$($Script:DefCannotChangePassword)" minOccurs="0"/>
            <xs:element type="xs:boolean" name="$($Script:DefChangePasswordAtLogon)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefDescription)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefDisplayName)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefDistinguishedName)"/>
            <xs:element type="xs:boolean" name="$($Script:DefEnabled)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefGivenName)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefInitials)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefName)"/>
            <xs:element type="xs:boolean" name="$($Script:DefPasswordNeverExpires)" minOccurs="0"/>
            <xs:element type="xs:boolean" name="$($Script:DefPasswordNotRequired)" minOccurs="0"/>
            <xs:element type="xs:boolean" name="$($Script:DefProtectedFromDel)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefSamAccountName)"/>
            <xs:element type="xs:boolean" name="$($Script:DefSmartcardLogonRequired)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefSurname)" minOccurs="0"/>
            <xs:element type="xs:boolean" name="$($Script:DefTrustedForDelegation)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefUserPrincipalName)" minOccurs="0"/>
            <xs:element type="permissionsType" name="$($Script:DefPermissions)" minOccurs="0"/>
        </xs:all>
    </xs:complexType>
    <xs:complexType name="usersType">
        <xs:sequence>
            <xs:element type="userType" name="$($Script:DefUser)" minOccurs="0" maxOccurs="unbounded"/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name="computerType">
        <xs:all>
            <xs:element type="xs:string" name="$($Script:DefDescription)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefDistinguishedName)"/>
            <xs:element type="xs:boolean" name="$($Script:DefProtectedFromDel)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefSamAccountName)"/>
            <xs:element type="xs:string" name="$($Script:DefName)"/>
            <xs:element type="permissionsType" name="$($Script:DefPermissions)" minOccurs="0"/>
        </xs:all>
    </xs:complexType>
    <xs:complexType name="computersType">
        <xs:sequence>
            <xs:element type="computerType" name="$($Script:DefComputer)" minOccurs="0" maxOccurs="unbounded"/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name="membersType">
        <xs:sequence>
            <xs:element type="wellKnownType" name="$($Script:DefMember)" minOccurs="0" maxOccurs="unbounded"/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name="groupType">
        <xs:all>
            <xs:element name="$($Script:DefCategory)">
                <xs:simpleType>
                    <xs:restriction base="xs:string">
                        <xs:enumeration value="Distribution"/>
                        <xs:enumeration value="Security"/>
                    </xs:restriction>
                </xs:simpleType>
            </xs:element>
            <xs:element type="xs:string" name="$($Script:DefDescription)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefDistinguishedName)"/>
            <xs:element type="xs:string" name="$($Script:DefDisplayName)" minOccurs="0"/>
            <xs:element type="membersType" name="$($Script:DefMembers)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefName)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefSamAccountName)"/>
            <xs:element name="$($Script:DefScope)">
                <xs:simpleType>
                    <xs:restriction base="xs:string">
                        <xs:enumeration value="DomainLocal"/>
                        <xs:enumeration value="Global"/>
                        <xs:enumeration value="Universal"/>
                    </xs:restriction>
                </xs:simpleType>
            </xs:element>
            <xs:element type="xs:boolean" name="$($Script:DefProtectedFromDel)" minOccurs="0"/>
            <xs:element type="permissionsType" name="$($Script:DefPermissions)" minOccurs="0"/>
        </xs:all>
    </xs:complexType>
    <xs:complexType name="groupsType">
        <xs:sequence>
            <xs:element type="groupType" name="$($Script:DefGroup)" minOccurs="0" maxOccurs="unbounded"/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name="wmiFilterType">
        <xs:all>
            <xs:element type="xs:string" name="$($Script:DefWmiParm1)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefWmiParm2)"/>
            <xs:element type="xs:string" name="$($Script:DefWmiAuthor)" minOccurs="0"/>
            <xs:element type="xs:boolean" name="$($Script:DefShowAdvView)" minOccurs="0"/>
            <xs:element type="xs:string" name="$($Script:DefWmiName)"/>
            <xs:element type="permissionsType" name="$($Script:DefPermissions)" minOccurs="0"/>
        </xs:all>
    </xs:complexType>
    <xs:complexType name="wmiFiltersType">
        <xs:sequence>
            <xs:element type="wmiFilterType" name="$($Script:DefWmiFilter)" minOccurs="0" maxOccurs="unbounded"/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name="linkType">
        <xs:all>
            <xs:element type="xs:boolean" name="$($Script:DefEnabled)" minOccurs="0"/>
            <xs:element type="xs:boolean" name="$($Script:DefEnforced)" minOccurs="0"/>
            <xs:element name="$($Script:DefOrder)">
                <xs:simpleType>
                    <xs:restriction base="xs:integer">
                        <xs:minInclusive value="0"/>
                    </xs:restriction>
                </xs:simpleType>
            </xs:element>
            <xs:element type="xs:string" name="$($Script:DefContainer)"/>
        </xs:all>
    </xs:complexType>
    <xs:complexType name="linksType">
        <xs:sequence>
            <xs:element type="linkType" name="$($Script:DefLink)" minOccurs="0" maxOccurs="unbounded"/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name="groupPolicyType">
        <xs:all>
            <xs:element type="xs:string" name="$($Script:DefName)"/>
            <xs:element type="xs:string" name="$($Script:DefTargetName)" minOccurs="0"/>
            <xs:element name="$($Script:DefGpoStatus)">
                <xs:simpleType>
                    <xs:restriction base="xs:string">
                        <xs:enumeration value="AllSettingsEnabled"/>
                        <xs:enumeration value="AllSettingsDisabled"/>
                        <xs:enumeration value="UserSettingsDisabled"/>
                        <xs:enumeration value="ComputerSettingsDisabled"/>
                    </xs:restriction>
                </xs:simpleType>
            </xs:element>
            <xs:element type="linksType" name="$($Script:DefLinks)" minOccurs="0"/>
            <xs:element type="permissionsType" name="$($Script:DefPermissions)" minOccurs="0"/>
        </xs:all>
    </xs:complexType>
    <xs:complexType name="groupPoliciesType">
        <xs:sequence>
            <xs:element type="groupPolicyType" name="$($Script:DefGroupPolicy)" minOccurs="0" maxOccurs="unbounded"/>
        </xs:sequence>
    </xs:complexType>
    <xs:complexType name="configurationType">
        <xs:all>
            <xs:element type="parametersType" name="$($Script:DefParameters)" minOccurs="0" maxOccurs="1"/>
            <xs:element type="organizationalUnitsType" name="$($Script:DefOus)" minOccurs="0" maxOccurs="1"/>
            <xs:element type="usersType" name="$($Script:DefUsers)" minOccurs="0" maxOccurs="1"/>
            <xs:element type="computersType" name="$($Script:DefComputers)" minOccurs="0" maxOccurs="1"/>
            <xs:element type="groupsType" name="$($Script:DefGroups)" minOccurs="0" maxOccurs="1"/>
            <xs:element type="wmiFiltersType" name="$($Script:DefWmiFilters)" minOccurs="0" maxOccurs="1"/>
            <xs:element type="groupPoliciesType" name="$($Script:DefGroupPolicies)" minOccurs="0" maxOccurs="1"/>
        </xs:all>
    </xs:complexType>
</xs:schema>
"@

# Minimum Windows version required to execute this script.
Set-Variable -Scope Script -Option Constant -Name WindowsMinVersion -Value ([Version]'6.3')

# Default Settings Xml file name.
Set-Variable -Scope Script -Option Constant -Name SettingsXmlFile -Value 'settings'

# Default GPOs Backup folder name.
Set-Variable -Scope Script -Option Constant -Name GpoBackupFolderName -Value 'Backup'

# Default ADMX folder name.
Set-Variable -Scope Script -Option Constant -Name AdmxFolderName -Value 'Admx'

# Default GPOs Report folder name.
Set-Variable -Scope Script -Option Constant -Name GpoReportFolderName -Value 'Reports'

# Default All GPOs Report folder name.
Set-Variable -Scope Script -Option Constant -Name GpoAllReportsFolderName -Value 'Group Policy Objects'

# Default Domain GPOs Report folder name.
Set-Variable -Scope Script -Option Constant -Name GpoDomainReportsFolderName -Value 'Domain'

# Default Migration Table name.
Set-Variable -Scope Script -Option Constant -Name MigrationTableFile -Value 'template.migtable'

# Temporary folder used by the script.
Set-Variable -Scope Script -Name TempFolder -Value $null

# Base folder for script's output files.
Set-Variable -Scope Script -Name BaseFolder -Value $null

# Content of the log file.
Set-Variable -Scope Script -Name LogContent -Value $null

# Cache to store already resolved names.
Set-Variable -Scope Script -Name NameToSidCache -Value @{}

# Active Directory Schema attributes.
Set-Variable -Scope Script -Name SchemaAtributes -Value @{}

# System Container Identifiers.
Set-Variable -Scope Script -Option Constant -Name SystemContainersId -Value @{
    'NTDSQuotas' = '6227F0AF1FC2410D8E3BB10615BB5B0F'
    'Microsoft' = 'F4BE92A4C777485E878E9421D53087DB'
    'ProgramData' = '09460C08AE1E4A4EA0F64AEE7DAA1E5A'
    'ForeignSecurityPrincipals' = '22B70C67D56E4EFB91E9300FCA3DC1AA'
    'DeletedObjects' = '18E2EA80684F11D2B9AA00C04F79F805'
    'Infrastructure' = '2FBAC1870ADE11D297C400C04FD8D5CD'
    'LostAndFound' = 'AB8153B7768811D1ADED00C04FD8D5CD'
    'System' = 'AB1D30F3768811D1ADED00C04FD8D5CD'
    'DomainControllers' = 'A361B2FFFFD211D1AA4B00C04FD7D83A'
    'Computers' = 'AA312825768811D1ADED00C04FD8D5CD'
    'Users' = 'A9D1CA15768811D1ADED00C04FD8D5CD'
    'ManagedServiceAccounts' = '1EB93889E40C45DF9F0C64D23BBB6237'
}

# List of well-known SIDs.
Set-Variable -Scope Script -Option Constant -Name WellKnownSids -Value @{
    'S-1-0' = 'Null Authority'
    'S-1-0-0' = 'Null Sid'
    'S-1-1' = 'World Authority'
    'S-1-1-0' = 'Everyone'
    'S-1-2' = 'Local Authority'
    'S-1-2-0' = 'Local'
    'S-1-2-1' = 'Console Logon'
    'S-1-3' = 'Creator Authority'
    'S-1-3-0' = 'Creator Owner'
    'S-1-3-1' = 'Creator Group'
    'S-1-3-2' = 'Creator Owner Server'
    'S-1-3-3' = 'Creator Group Server'
    'S-1-3-4' = 'Owner Rights'
    'S-1-4' = 'Non-unique Authority'
    'S-1-5' = 'NT Pseudo Domain'
    'S-1-5-1' = 'NT Authority\Dialup'
    'S-1-5-2' = 'NT Authority\Network'
    'S-1-5-3' = 'NT Authority\Batch'
    'S-1-5-4' = 'NT Authority\Interactive'
    'S-1-5-6' = 'NT Authority\Service'
    'S-1-5-7' = 'NT Authority\Anonymous Logon'
    'S-1-5-8' = 'NT Authority\Proxy'
    'S-1-5-9' = 'NT Authority\Enterprise Domain Controllers'
    'S-1-5-10' = 'NT Authority\Self'
    'S-1-5-11' = 'NT Authority\Authenticated Users'
    'S-1-5-12' = 'NT Authority\Restricted'
    'S-1-5-13' = 'NT Authority\Terminal Server Users'
    'S-1-5-14' = 'NT Authority\Remote Interactive Logon'
    'S-1-5-15' = 'NT Authority\This Organization'
    'S-1-5-17' = 'IUSR'
    'S-1-5-18' = 'NT Authority\System'
    'S-1-5-19' = 'NT Authority\Local Service'
    'S-1-5-20' = 'NT Authority\Network Service'
    'S-1-5-33' = 'NT Authority\Write Restricted'
    'S-1-5-1000' = 'NT Authority\Other Organization'
    'S-1-5-21-498' = 'Enterprise Read-only Domain Controllers'
    'S-1-5-21-500' = 'Administrator'
    'S-1-5-21-501' = 'Guest'
    'S-1-5-21-502' = 'krbtgt'
    'S-1-5-21-512' = 'Domain Admins'
    'S-1-5-21-513' = 'Domain Users'
    'S-1-5-21-514' = 'Domain Guests'
    'S-1-5-21-515' = 'Domain Computers'
    'S-1-5-21-516' = 'Domain Controllers'
    'S-1-5-21-517' = 'Cert Publishers'
    'S-1-5-21-518' = 'Schema Admins'
    'S-1-5-21-519' = 'Enterprise Admins'
    'S-1-5-21-520' = 'Group Policy Creator Owners'
    'S-1-5-21-521' = 'Read-only Domain Controllers'
    'S-1-5-21-522' = 'Cloneable Domain Controllers'
    'S-1-5-21-526' = 'Key Admins'
    'S-1-5-21-527' = 'Enterprise Key Admins'
    'S-1-5-21-553' = 'RAS and IAS Servers'
    'S-1-5-21-571' = 'Allowed RODC Password Replication Group'
    'S-1-5-21-572' = 'Denied RODC Password Replication Group'
    'S-1-5-21-1101' = 'DnsAdmins'
    'S-1-5-32' = 'Builtin'
    'S-1-5-32-544' = 'Administrators'
    'S-1-5-32-545' = 'Users'
    'S-1-5-32-546' = 'Guests'
    'S-1-5-32-547' = 'Power Users'
    'S-1-5-32-548' = 'Account Operators'
    'S-1-5-32-549' = 'Server Operators'
    'S-1-5-32-550' = 'Print Operators'
    'S-1-5-32-551' = 'Backup Operators'
    'S-1-5-32-552' = 'Replicators'
    'S-1-5-32-554' = 'Builtin\Pre-Windows 2000 Compatible Access'
    'S-1-5-32-555' = 'Builtin\Remote Desktop Users'
    'S-1-5-32-556' = 'Builtin\Network Configuration Operators'
    'S-1-5-32-557' = 'Builtin\Incoming Forest Trust Builders'
    'S-1-5-32-558' = 'Builtin\Performance Monitor Users'
    'S-1-5-32-559' = 'Builtin\Performance Log Users'
    'S-1-5-32-560' = 'Builtin\Windows Authorization Access Group'
    'S-1-5-32-561' = 'Builtin\Terminal Server License Servers'
    'S-1-5-32-562' = 'Builtin\Distributed COM Users'
    'S-1-5-32-568' = 'Builtin\IIS_IUSRS'
    'S-1-5-32-569' = 'Builtin\Cryptographic Operators'
    'S-1-5-32-573' = 'Builtin\Event Log Readers'
    'S-1-5-32-574' = 'Builtin\Certificate Service DCOM Access'
    'S-1-5-32-575' = 'Builtin\RDS Remote Access Servers'
    'S-1-5-32-576' = 'Builtin\RDS Endpoint Servers'
    'S-1-5-32-577' = 'Builtin\RDS Management Servers'
    'S-1-5-32-578' = 'Builtin\Hyper-V Administrators'
    'S-1-5-32-579' = 'Builtin\Access Control Assistance Operators'
    'S-1-5-32-580' = 'Builtin\Remote Management Users'
    'S-1-5-32-582' = 'Builtin\Storage Replica Administrators'
    'S-1-5-64-10' = 'NT Authority\NTLM Authentication'
    'S-1-5-64-14' = 'NT Authority\Channel Authentication'
    'S-1-5-64-21' = 'NT Authority\Digest Authority'
    'S-1-5-80' = 'NT Service'
    'S-1-5-80-0' = 'All Services'
    'S-1-5-80-2387347252-3645287876-2469496166-3824418187-3586569773' = 'NT SERVICE\ALG'
    'S-1-5-80-4059739203-877974739-1245631912-527174227-2996563517' = 'NT SERVICE\Wecsvc'
    'S-1-5-80-569256582-2953403351-2909559716-1301513147-412116970' = 'NT SERVICE\WinRM'
    'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464' = 'NT SERVICE\TrustedInstaller'
    'S-1-5-83-0' = 'NT Virtual Machine\Virtual Machines'
    'S-1-5-90-0' = 'Windows Manager\Windows Manager Group'
    'S-1-5-113' = 'Local account'
    'S-1-5-114' = 'Local account and member of Administrators group'
    'S-1-7' = 'Internet$'
    'S-1-16' = 'Mandatory Label'
    'S-1-16-0' = 'Mandatory Label\Untrusted Mandatory Level'
    'S-1-16-4096' = 'Mandatory Label\Low Mandatory Level'
    'S-1-16-8192' = 'Mandatory Label\Medium Mandatory Level'
    'S-1-16-8448' = 'Mandatory Label\Medium Plus Mandatory Level'
    'S-1-16-12288' = 'Mandatory Label\High Mandatory Level'
    'S-1-16-16384' = 'Mandatory Label\System Mandatory Level'
    'S-1-16-20480' = 'Mandatory Label\Protected Process Mandatory Level'
    'S-1-16-28672' = 'Mandatory Label\Secure Process Mandatory Level'
}

#endregion

#region Generic Helpers

if (!(Get-Command -Name 'Add-Log' -CommandType Function -ErrorAction SilentlyContinue)) {
    function Add-Log {
    <#
    .SYNOPSIS
        Logs messages.
    .DESCRIPTION
        Logs messages.
    .PARAMETER Log
        Message to display.
    .PARAMETER Type
        Type of message.
    .PARAMETER NoNewline
        If specified, the display continues on the same line.
    .OUTPUTS
        None.
    #>
        param (
            [Parameter(Mandatory=$true)]
            [String]$Log = $null,

            [Parameter(Mandatory=$false)]
            [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Title1', 'Title2')]
            [String]$Type = 'Info',

            [Parameter(Mandatory=$false)]
            [Switch]$NoNewline = $false
        )
        $color = [System.ConsoleColor]::DarkYellow
        $colors = @{
            'info' = [System.ConsoleColor]::DarkGray
            'error' = [System.ConsoleColor]::Red
            'warning' = [System.ConsoleColor]::Yellow
            'success' = [System.ConsoleColor]::Green
            'title1' = [System.ConsoleColor]::Magenta
            'title2' = [System.ConsoleColor]::Cyan
        }
        if ([String]::IsNullOrEmpty($Type)) {
            $Type = 'info'
        } else {
            $Type = $Type.ToLower()
        }
        if ($colors.ContainsKey($Type)) {
            $color = $colors[$Type]
        }
        $t = 1
        if ($Type -ieq 'warning') {
            $t = 2
        } elseif ($Type -ieq 'error') {
            $t = 3
        }
        $date = Get-Date -ErrorAction SilentlyContinue
        $time = '{0:HH:mm:ss.ffff}' -f $date
        $day = '{0:yyyy-MM-dd}' -f $date
        $component = $null
        $context = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $thread = [Threading.Thread]::CurrentThread.ManagedThreadId
        $file = $null
        $line = $null
        try {
            $callstack = Get-PSCallStack -ErrorAction Stop
            if (($callstack -is [Array]) -and ($callstack.Count -gt 0)) {
                $callstack = $callstack[1]
                $component = $callstack.FunctionName
                if (![String]::IsNullOrEmpty($callstack.ScriptName)) {
                    $line = $callstack.ScriptLineNumber
                    $file = '{0}:{1}' -f $callstack.ScriptName, $line
                }
            }
        } catch {
            Write-Debug "Unable to list the call stack. $_" -ErrorAction SilentlyContinue
        }
        if ((Test-Path -Path 'Variable:Script:LogContent')) {
            if (!$Script:LogContent) {
                $Script:LogContent = @()
            }
            $Script:LogContent += [PSCustomObject]@{
                'Day' = $day
                'Time' = $time
                'Log' = $Log
                'Component' = $component
                'Context' = $context
                'Type' = $t
                'Thread' = $thread
                'File' = $file
            }
        }
        try {
            if (($Type -ieq 'Error') -and ![String]::IsNullOrEmpty($line)) {
                $Log += " (line $line)"
            }
            if ($true) {
                Write-Host -Object $Log -ForegroundColor $color -NoNewline:$NoNewline -ErrorAction Stop
            } else {
                Write-Output -InputObject $Log -ErrorAction Stop
            }
        } catch {
            Write-Debug "Unable to write to the host/standard output. $_" -ErrorAction SilentlyContinue
        }
    }
}

function Get-CurrentWindowsVersion {
<#
.SYNOPSIS
    Gets current version of Windows.
.DESCRIPTION
    Gets current version of Windows.
.OUTPUTS
    Version number or null.
#>
    $version = $null
    try {
        $osInfo = Get-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $major = [Int32]$osInfo.GetValue('CurrentMajorVersionNumber')
        $minor = [Int32]$osInfo.GetValue('CurrentMinorVersionNumber')
        if ($major -iin @($null, 0)) {
            $major, $minor = $osInfo.GetValue('CurrentVersion').ToString().Split('.')
        }
        $build = [Int32]$osInfo.GetValue('CurrentBuildNumber')
        $revision = [Int32]$osInfo.GetValue('UBR')
        $version = [Version](($major, $minor, $build, $revision) -join '.')
        Add-Log -Log "Windows version: $($version)"
    } catch {
        Add-Log -Log "Can not get Windows version. $_" -Type Error
        return $null
    }
    return $version
}

function Test-Admin {
<#
.SYNOPSIS
    Checks if the current Windows PowerShell session is running as Administrator.
.DESCRIPTION
    Checks if the current Windows PowerShell session is running as Administrator.
.OUTPUTS
    True if Admin, false otherwise.
#>
    return ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
                [Security.Principal.WindowsBuiltInRole]::Administrator)

}

function Get-ComplexPassword {
<#
.SYNOPSIS
    Generates a complex password.
.DESCRIPTION
    Generates a complex password.
.PARAMETER Length
    Password's length.
.PARAMETER CharMaxOccurrence
    Maximum occurrences of a character.
.PARAMETER MaxIteration
    Maximum number of iterations for this function.
.PARAMETER ClearTextPassword
    Return a clear text password instead of a Secure String.
.OUTPUTS
    Complex password or null.
#>
    param(
        [parameter(mandatory=$false)]
        [UInt16]$Length = 25,

        [parameter(mandatory=$false)]
        [UInt16]$CharMaxOccurrence = 4,

        [parameter(mandatory=$false)]
        [UInt16]$MaxIteration = 100,

        [parameter(mandatory=$false)]
        [Switch]$ClearTextPassword = $false
    )
    $lettersLC = 'abcdefghijklmnopqrstuvwxyz'
    $lettersUC = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $digits = '0123456789'
    $symbols = '^~!@#$%&*_+=`|\(){}[]:;"''<->,.?/'
    $minLength = 8
    $allowedChars = $lettersLC + $lettersUC + $digits + $symbols
    if ($Length -lt $minLength) {
        Add-Log -Log "Password length must be at least $($minLength) characters" -Type Error
        return $null
    }
    if ($Length -gt 256) {
        Add-Log -Log 'Password length must not exceed 256 characters' -Type Error
        return $null
    }
    if ($CharMaxOccurrence -le 0) {
        $CharMaxOccurrence = 1
    }
    $bytes = New-Object -TypeName 'System.Byte[]' -ArgumentList $Length
    $rnd = New-Object -TypeName System.Security.Cryptography.RNGCryptoServiceProvider
    $rnd.GetBytes($bytes)
    $p = ''
    for ($i = 0; $i -lt $Length; $i++) {
        $p += $allowedChars[$bytes[$i] % $allowedChars.Length]
    }
    $isComplex = $true
    $array = $p.ToCharArray()
    $count = $array | Group-Object -NoElement | Where-Object { $_.Count -gt $CharMaxOccurrence }
    if ($count) {
        if ($MaxIteration -gt 1) {
            $PSBoundParameters['MaxIteration'] = $MaxIteration - 1
            return (Get-ComplexPassword @PSBoundParameters)
        } else {
            $isComplex = $false
        }
    }
    if ($isComplex) {
        foreach ($criteria in @($lettersLC, $lettersUC, $digits, $symbols)) {
            $count = $array | Where-Object { $_ -cin $criteria.ToCharArray() }
            if (!$count) {
                if ($MaxIteration -gt 1) {
                    $PSBoundParameters['MaxIteration'] = $MaxIteration - 1
                    return (Get-ComplexPassword @PSBoundParameters)
                } else {
                    $isComplex = $false
                    break
                }
            }
        }
    }
    if (!$isComplex) {
        Add-Log -Log 'Generated password doesn''t meet the complexity requirements' -Type Warning
    }
    if (!$ClearTextPassword) {
        try {
            $p = ConvertTo-SecureString -String $p -AsPlainText -Force -ErrorAction Stop
        } catch {
            Add-Log -Log "Unable to build a Secure String. $_" -Type Error
            return $null
        }
    }
    return $p
}

function Get-Confirmation {
<#
.SYNOPSIS
    Asks for confirmation.
.DESCRIPTION
    Asks for confirmation.
.PARAMETER Message
    Message to display.
.PARAMETER Answers
    Accepted answers.
.PARAMETER Loop
    If true, loop until a valid answer is provided.
.OUTPUTS
    True if the user wants to continue, false otherwise.
#>
    param (
        [Parameter(Mandatory=$false)]
        [String]$Message = 'Do you want to continue: [Y]es or [N]o?',

        [Parameter(Mandatory=$false)]
        [String[]]$Answers = @('y', 'yes'),

        [Parameter(Mandatory=$false)]
        [Switch]$Loop = $false
    )
    try {
        do {
            $answer = Read-Host -Prompt $Message -ErrorAction Stop
            if ($answer -iin $Answers) {
                return $true
            }
        } while ($Loop)
    } catch {
        Add-Log -Log "An error occurred while asking for confirmation. $_" -Type Error
    }
    return $false
}

function Get-LogFile {
<#
.SYNOPSIS
    Builds the log file.
.DESCRIPTION
    Builds the log file.
.PARAMETER File
    Path to the log file.
.PARAMETER Format
    Format of the log file.
.OUTPUTS
    True on scucess, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [String]$File = $null,

        [Parameter(Mandatory=$false)]
        [ValidateSet('CMTrace', 'Csv', 'Html', 'Json', 'Xml')]
        [String]$Format = 'CMTrace',

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    if ([String]::IsNullOrEmpty($File) -or
        !(Test-Path -Path $File -IsValid) -or
        !(Test-Path -Path 'Variable:Script:LogContent')) {
        return $false
    }
    $parameters = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    $folder = Split-Path -Path $File -Parent @parameters
    if ([String]::IsNullOrEmpty($folder)) {
        $File = Join-Path -Path '.' -ChildPath $File
    }
    $content = ''
    switch ($Format) {
        'CMTrace' {
            foreach ($e in $Script:LogContent) {
                if (!$e) {
                    continue
                }
                $entry = 'time="{0}" date="{1}" component="{2}" context="{3}" type="{4}" thread="{5}" file="{6}"'
                $entry = $entry -f $e.Time, $e.Day, $e.Component, $e.Context, $e.Type, $e.Thread, $e.File
                $entry = "<![LOG[$($e.Log)]LOG]!><$($entry)>" + [Environment]::NewLine
                $content += $entry
            }
            break
        }
        'Csv' {
            $content = $Script:LogContent | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation @parameters
            break
        }
        'Html' {
            $content = $Script:LogContent | ConvertTo-Html -As Table @parameters
            break
        }
        'Json' {
            $content = $Script:LogContent | ConvertTo-Json -Compress @parameters
            break
        }
        'Xml' {
            $content = $Script:LogContent | ConvertTo-Xml -As String -NoTypeInformation @parameters
            break
        }
        default {
            Add-Log "Log format '$Format' is not supported" -Type Error
            return $false
        }
    }
    try {
        $content | Out-File -FilePath $File -Encoding 'utf8' -Force:$Force @parameters
        Add-Log -Log "Log file generated in '$File'"
    } catch {
        Add-Log -Log "Unable to write to the file '$File'. $_" -Type Error
        return $false
    }
    return $true
}

function Get-SanitizedFileName {
<#
.SYNOPSIS
    Sanitizes a file's name.
.DESCRIPTION
    Sanitizes a file's name.
.PARAMETER Name
    Name to sanitize.
.PARAMETER MaxSize
    Maximum name's size.
.OUTPUTS
    Sanitized name.
#>
    param (
        [Parameter(Mandatory=$true)]
        [String]$Name = $null,

        [Parameter(Mandatory=$false)]
        [UInt16]$MaxSize = 260
    )
    $disallowedCharacters = @( '\', '/', ':', '*', '?', '"', '<', '>', '|')
    $chars = $Name.ToCharArray()
    $removed = ''
    $truncated = ''
    foreach ($char in $chars) {
        $b = [Byte][Char]$char
        $c = $char
        $len = 1
        if (($b -lt 0x20) -or ($char -iin $disallowedCharacters)) {
            $c = '!00' + ('{0:x2}' -f $b)
            $len = $c.Length
        }
        if (($truncated.Length + $len) -le $MaxSize) {
            $truncated += $c
        } else {
            $removed += $c
        }
    }
    if ([String]::IsNullOrEmpty($removed)) {
        return $truncated
    }
    $chars = $removed.ToCharArray()
    [UInt16]$h = 0
    foreach ($char in $chars) {
        [UInt16]$lowBit = 0
        if (0x8000 -band $h) {
            $lowBit = 1
        }
        $h = [UInt16](($h -shl 1) -bor $lowBit) + [UInt16]$char
    }
    $h = "$h".PadLeft(5, '0')
    $n = "$($truncated)-$($h)"
    if ($n.Length -gt $MaxSize) {
        return $truncated
    }
    return $n
}

function Get-StringHash {
<#
.SYNOPSIS
    Computes a hash from a string or a bytes array.
.DESCRIPTION
    Computes a hash from a string or a bytes array.
.PARAMETER Object
    String or bytes array.
.PARAMETER Algorithm
    Hash algorithm.
.PARAMETER Salt
    Salt added at the end.
.OUTPUTS
    Hash on success, null otherwise.
#>
    param (
        [Parameter(Mandatory=$true)]
        [Object]$Object = $null,

        [Parameter(Mandatory=$false)]
        [ValidateSet('MD5', 'SHA1', 'SHA256', 'SHA384', 'SHA512')]
        [String]$Algorithm = 'SHA256',

        [Parameter(Mandatory=$false)]
        [Object]$Salt = $null
    )
    try {
        $enc = [System.Text.Encoding]::UTF8
        $data = $Object
        if ($data -is [String]) {
            $data = $enc.GetBytes($data)
        }
        if ($data -isnot [Byte[]]) {
            return $null
        }
        if ($Salt) {
            if ($Salt -is [String]) {
                $Salt = $enc.GetBytes($Salt)
            }
            if ($Salt -isnot [Byte[]]) {
                return $null
            }
            $data += $Salt
        }
        $hash = [Security.Cryptography.HashAlgorithm]::Create($Algorithm)
        $bytes = $hash.ComputeHash($data)
        $hashString = ''
        $bytes | ForEach-Object { $hashString += ('{0:x2}' -f $_).PadLeft(2, '0') }
        return $hashString
    } catch {
        Add-Log -Log "Unable to compute hash. $_" -Type Error
    }
    return $null
}

#endregion

#region AD Helpers

function Test-ProtectedFromDeletion {
<#
.SYNOPSIS
    Checks if an object is protected from accidental deletion.
.DESCRIPTION
    Checks if an object is protected from accidental deletion.
.PARAMETER Rules
    ACLs list.
.OUTPUTS
    True if the object is protected, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Security.AccessControl.AuthorizationRule[]]$Rules = $null
    )
    if (!$Rules) {
        Add-Log -Log 'ACL Rule must be valid' -Type Error
        return $false
    }
    $everyone = 'S-1-1-0'
    $rights = [System.DirectoryServices.ActiveDirectoryRights]::DeleteTree -band
              [System.DirectoryServices.ActiveDirectoryRights]::Delete
    $acl = $null
    try {
        $acl = $Rules | Where-Object {
            ($_.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Deny) -and
            (($_.ActiveDirectoryRights -band $rights) -eq $rights) -and
            ($_.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value -ieq $everyone)
        }
    } catch {
        Add-Log -Log "Error while reading ACL rule. $_" -Type Warning
        $acl = $null
    }
    return ($acl -ne $null)
}

function Test-WellKnownSid {
<#
.SYNOPSIS
    Checks if a SID is a well known SID.
.DESCRIPTION
    Checks if a SID is a well known SID.
.PARAMETER Sid
    SID to check.
.OUTPUTS
    True if the SID is a well known SID, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [Security.Principal.SecurityIdentifier]$Sid = $null
    )
    $wkSids = [System.Enum]::GetValues([System.Security.Principal.WellKnownSidType])
    foreach ($wkSid in $wkSids) {
        if ($Sid.IsWellKnown($wkSid)) {
            return $true
        }
    }
    $s = "$Sid".ToUpper()
    if ($s -imatch '^(S-1-5-21-\d+-\d+-\d+)-(\d+)$') {
        $s = "S-1-5-21-$($Matches[2])"
    }
    return $Script:WellKnownSids.ContainsKey($s)
}

function Convert-StringToSid {
<#
.SYNOPSIS
    Converts a String to Sid.
.DESCRIPTION
    Converts a String to Sid.
.PARAMETER String
    String to convert.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    Sid object or null otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [String]$String = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $sid = $null
    try {
        $sid = New-Object -TypeName System.Security.Principal.SecurityIdentifier($String)
    } catch {
        $sid = $null
    }
    if ($sid -is [System.Security.Principal.SecurityIdentifier]) {
        return $sid
    }
    $str = $String.ToLower()
    if ($Script:NameToSidCache.ContainsKey($str)) {
        return $Script:NameToSidCache[$str]
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        'Server' = $Server
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    $found = $null
    try {
        $s = $String
        $dom = $null
        if ($s.Contains('\')) {
            $dom, $s = $s.Split('\')
        }
        if ($s.Contains('@')) {
            $s, $dom = $s.Split('@')
        }
        if ($Domain -and ($dom -iin @($Domain.NetBIOSName, $Domain.RootNetBIOSName))) {
            if ([String]::IsNullOrEmpty($commonParams['Server'])) {
                $commonParams['Server'] = $Domain.PDCEmulator
            }
            if ($dom -ieq $Domain.RootNetBIOSName) {
                $commonParams['Server'] = $Domain.RootPDCEmulator
            }
            $found = Get-ADObject -Filter "samAccountName -eq '$s'" -Properties objectSid @commonParams
        }
    } catch {
        if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
            Add-Log -Log "Error while getting object '$s' information. $_" -Type Error
            return $null
        }
    }
    if ($found -is [Microsoft.ActiveDirectory.Management.ADObject]) {
        $sid = $found.ObjectSid
    } else {
        try {
            $account = New-Object -TypeName System.Security.Principal.NTAccount($String)
            $sid = $account.Translate([System.Security.Principal.SecurityIdentifier])
        } catch {
            return $null
        }
    }
    if ($sid -is [System.Security.Principal.SecurityIdentifier]) {
        $Script:NameToSidCache.Add($str, $sid)
    }
    return $sid
}

function Convert-SidToString {
<#
.SYNOPSIS
    Resolves SID.
.DESCRIPTION
    Resolves SID.
.PARAMETER SID
    SID to resolve.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    Account name or null otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Security.Principal.SecurityIdentifier]$Sid = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    foreach ($name in $Script:NameToSidCache.Keys) {
        if ($Script:NameToSidCache[$name] -ieq $Sid) {
            return $name
        }
    }
    $parameters = @{
        'Properties' = @('objectSid', 'SamAccountName')
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        'Server' = $Domain.PDCEmulator
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters['Server'] = $Server
    }
    if ($Credential) {
        $parameters.Add('Credential', $Credential)
    }
    $sam = $null
    try {
        $obj = Get-ADObject -Filter "objectSid -eq '$($Sid)'" @parameters
        if ($obj) {
            $sam = "$($obj.SamAccountName)"
        }
    } catch {
        $sam = $null
    }
    if ([String]::IsNullOrEmpty($sam)) {
        try {
            $d = $Sid.Translate([System.Security.Principal.NTAccount])
            if ($d -and ![String]::IsNullOrEmpty($d.Value) -and $d.Value.Contains('\')) {
                $sam = $d.Value.Split('\')[1]
            }
        } catch {
            return $null
        }
    }
    if (!$Script:NameToSidCache.ContainsKey($sam)) {
        $Script:NameToSidCache.Add($sam, $Sid)
    }
    return $sam
}

function Get-SchemaAttributesFromXml {
<#
.SYNOPSIS
    Gets Active Directory Schema attributes.
.DESCRIPTION
    Gets Active Directory Schema attributes.
.PARAMETER Xml
    Xml document.
.PARAMETER Path
    Path to Xml file.
.PARAMETER Server
    Domain Controller to query.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    [CmdletBinding(DefaultParameterSetName='Xml')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='Xml')]
        [System.Xml.XmlDocument]$Xml = $null,

        [Parameter(Mandatory=$true, ParameterSetName='Path')]
        [String]$Path = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if ($Credential) {
        $parameters.Add('Credential', $Credential)
    }
    $Script:SchemaAtributes = @{}
    try {
        $rootDse = Get-ADRootDSE @parameters
        if (![String]::IsNullOrEmpty($Path) -and (Test-Path -Path $Path)) {
            [Xml]$Xml = Get-Content -Path $Path -ErrorAction Stop
        }
        $nodes = @() + (Select-Xml -Xml $Xml -XPath "//$($Script:DefObject)")
        $nodes += @(Select-Xml -Xml $Xml -XPath "//$($Script:DefInheritedObject)")
        $nodes = $nodes | Select-Object -Unique
        $filterSchemaId = ''
        $filterRights = ''
        if ($nodes) {
            foreach ($node in $nodes.Node.'#text') {
                if (!$node) {
                    continue
                }
                $guid = $null
                try {
                    $guid = [System.Guid]::Parse($node)
                } catch {
                    $guid = $null
                }
                if (!$guid) {
                    $filterSchemaId += "(lDAPDisplayName=$($node))"
                    $filterRights += "(DisplayName=$($node))"
                } else {
                    $hex = ($guid).ToByteArray()
                    $hex = '\' + (($hex | ForEach-Object -MemberName 'ToString' -ArgumentList 'X2') -join '\')
                    $filterSchemaId += "(schemaIDGUID=$($hex))"
                    $filterRights += "(rightsGuid=$($node))"
                }
            }
            $filter = "(|$($filterSchemaId))"
            $objs = Get-ADObject -SearchBase $rootDse.schemaNamingContext `
                                 -LDAPFilter $filter `
                                 -Properties 'lDAPDisplayName', 'schemaIDGUID' `
                                 @parameters
            foreach ($obj in $objs) {
                $guid = ([System.Guid]$obj.schemaIDGUID).ToString().ToLower()
                if (!$Script:SchemaAtributes.ContainsKey($guid)) {
                    $Script:SchemaAtributes.Add($guid, "$($obj.lDAPDisplayName)")
                }
            }
            $filter = "(|$($filterRights))"
            $objs = Get-ADObject -SearchBase "CN=Extended-Rights,$($rootDse.configurationNamingContext)" `
                                 -LDAPFilter $filter `
                                 -Properties 'RightsGuid', 'DisplayName' `
                                 @parameters
            foreach ($obj in $objs) {
                $guid = ([System.Guid]$obj.RightsGuid).ToString().ToLower()
                if (!$Script:SchemaAtributes.ContainsKey($guid)) {
                    $Script:SchemaAtributes.Add($guid, "$($obj.DisplayName)")
                }
            }
        }
    } catch {
        Add-Log -Log "Unable to get Active Directory Schema information. $_" -Type Error
        return $false
    }
    return $true
}

function Set-SchemaAttributesFromXml {
<#
.SYNOPSIS
    Sets Active Directory Schema attributes in Xml.
.DESCRIPTION
    Sets Active Directory Schema attributes in Xml.
.PARAMETER Xml
    Xml Document.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$Xml = $null
    )
    try {
        $nodes = @() + (Select-Xml -Xml $Xml -XPath "//$($Script:DefObject)")
        $nodes += @(Select-Xml -Xml $Xml -XPath "//$($Script:DefInheritedObject)")
        if ($nodes) {
            foreach ($node in $nodes.Node) {
                if (!$node) {
                    continue
                }
                $guid = $null
                try {
                    $guid = ([System.Guid]$node.InnerText).ToString().ToLower()
                    if ($Script:SchemaAtributes.ContainsKey($guid)) {
                        $node.InnerText = "$($Script:SchemaAtributes[$guid])"
                    }
                } catch {
                    Add-Log -Log "Unable to convert '$($node)' to Guid. $_" -Type Warning
                }
            }
        }
    } catch {
        Add-Log -Log "Unable to set Active Directory Schema information in Xml. $_" -Type Error
        return $false
    }
    return $true
}

function Get-AclFromXml {
<#
.SYNOPSIS
    Converts Xml nodes to Active Directory rights.
.DESCRIPTION
    Converts Xml nodes to Active Directory rights.
.PARAMETER Path
    Full path to Active Directory object.
.PARAMETER XmlNode
    Xml node.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    Custom object or null.
#>
    param(
        [Parameter(Mandatory=$true)]
        [String]$Path = $null,

        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$XmlNode = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $res = New-Object -TypeName System.Object
    $res | Add-Member -MemberType NoteProperty -Name 'Path' -Value $Path -Force
    $res | Add-Member -MemberType NoteProperty -Name 'AceList' -Value @() -Force
    $success = $true
    foreach ($x in $XmlNode.ChildNodes) {
        try {
            $access = Get-XmlValue -Xml $x -XmlPath $Script:DefAccessControl
            if ($access -ne $null) {
                $access = [System.Security.AccessControl.AccessControlType]::$access
            } else {
                $access = [System.Security.AccessControl.AccessControlType]::Allow
            }
            $rights = 0
            foreach ($r in ((Get-XmlValue -Xml $x -XmlPath $Script:DefRights) -split ' ')) {
                $rights = $rights -bor [System.DirectoryServices.ActiveDirectoryRights]::$r
            }
            if ([Int]$rights -eq 0) {
                Add-Log -Log 'Active Directory Rights cannot be null (skipped)' -Type Warning
                continue
            }
            $inheritance = Get-XmlValue -Xml $x -XmlPath $Script:DefInheritance
            if ($inheritance -ne $null) {
                $inheritance = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::$inheritance
            } else {
                $inheritance = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::None
            }
            $obj = Get-XmlValue -Xml $x -XmlPath $Script:DefObject
            if ($obj -ne $null) {
                try {
                    $obj = [System.Guid]$obj
                } catch {
                    $found = $false
                    foreach ($guid in $Script:SchemaAtributes.Keys) {
                        if ($Script:SchemaAtributes[$guid] -ieq "$obj") {
                            $obj = [System.Guid]$guid
                            $found = $true
                            break
                        }
                    }
                    if (!$found) {
                        throw "Unable to find Schema attribute '$($obj)' (Object)"
                    }
                }
            } else {
                $obj = [System.Guid]::Empty
            }
            $inObj = Get-XmlValue -Xml $x -XmlPath $Script:DefInheritedObject
            if ($inObj -ne $null) {
                try {
                    $inObj = [System.Guid]$inObj
                } catch {
                    $found = $false
                    foreach ($guid in $Script:SchemaAtributes.Keys) {
                        if ($Script:SchemaAtributes[$guid] -ieq "$inObj") {
                            $inObj = [System.Guid]$guid
                            $found = $true
                            break
                        }
                    }
                    if (!$found) {
                        throw "Unable to find Schema attribute '$($inObj)' (Inherited Object)"
                    }
                }
            } else {
                $inObj = [System.Guid]::Empty
            }
            $idRef = Get-XmlValue -Xml $x -XmlPath $Script:DefIdentityReference
            $isWellKnown = ((Get-XmlValue -Xml $idRef -XmlPath $Script:DefIsWellKnown) -ieq 'true')
            $idRef = Get-XmlValue -Xml $idRef -XmlPath '#text'
            if ($idRef -imatch '^(\[(Remote|Root)?DomainSID\])-(\d+)$') {
                if ($Matches[1] -ieq '[DomainSID]') {
                    $idRef = "$($Domain.DomainSID)-$($Matches[3])"
                } elseif ($Matches[1] -ieq '[RootDomainSID]') {
                    $idRef = "$($Domain.RootDomainSID)-$($Matches[3])"
                } else {
                    Add-Log -Log "Identity '$idRef' is from a remote domain (skipped)" -Type Warning
                    continue
                }
            }
            $sid = Convert-StringToSid -String $idRef -Server $Server -Credential $Credential -Domain $Domain
            if ($sid -isnot [System.Security.Principal.SecurityIdentifier]) {
                Add-Log -Log "Identity '$idRef' cannot be resolved (skipped)" -Type Warning
                continue
            }
            $res.AceList += New-Object -TypeName System.DirectoryServices.ActiveDirectoryAccessRule(
                $sid,
                $rights,
                $access,
                $obj,
                $inheritance,
                $inObj
            )
        } catch {
            Add-Log "Error while building ACE list. $_" -Type Warning
            $success = $false
        }
    }
    if (!$success) {
        return $null
    }
    return $res
}

function Test-AclInList {
<#
.SYNOPSIS
    Checks if ACLs are in a list.
.DESCRIPTION
    Checks if ACLs are in a list.
.PARAMETER Challengers
    Permissions to look for in References.
.PARAMETER References
    List of permissions.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True if Challengers are in References, False otherwise.
#>
    param(
        [Parameter(Mandatory=$false)]
        [System.Security.AccessControl.AuthorizationRule[]]$Challengers = $null,

        [Parameter(Mandatory=$false)]
        [System.Security.AccessControl.AuthorizationRule[]]$References = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    if (!$Challengers -and !$References) {
        return $true
    }
    if (!$Challengers -or !$References) {
        return $false
    }
    $commonParams = @{
        'Domain' = $Domain
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    foreach ($r in $Challengers) {
        if (!$r -or !$r.IdentityReference) {
            continue
        }
        $isIn = $false
        foreach ($ref in $References) {
            if (!$ref -or !$ref.IdentityReference) {
                continue
            }
            $b = ($r.ActiveDirectoryRights -ieq $ref.ActiveDirectoryRights) -and
                 ($r.InheritanceType -ieq $ref.InheritanceType) -and
                 ($r.ObjectType -ieq $ref.ObjectType) -and
                 ($r.InheritedObjectType -ieq $ref.InheritedObjectType) -and
                 ($r.ObjectFlags -ieq $ref.ObjectFlags) -and
                 ($r.AccessControlType -ieq $ref.AccessControlType) -and
                 ($r.InheritanceFlags -ieq $ref.InheritanceFlags) -and
                 ($r.PropagationFlags -ieq $ref.PropagationFlags)
            if (!$b) {
                continue
            }
            $sid = Convert-StringToSid -String $r.IdentityReference.Value @commonParams
            $sidRef = Convert-StringToSid -String $ref.IdentityReference.Value @commonParams
            if (($sid -and $sidRef -and ($sid -eq $sidRef)) -or
                (!$sid -and !$sidRef -and ($r.IdentityReference.Value -ieq $ref.IdentityReference.Value))) {
                $isIn = $true
                break
            }
        }
        if (!$isIn) {
            return $false
        }
    }
    return $true
}

function Get-XGPO {
<#
.SYNOPSIS
    Gets one GPO or all the GPOs in a domain.
.DESCRIPTION
    Gets one GPO or all the GPOs in a domain.
.PARAMETER Guid
    Guid of the GPO.
.PARAMETER Name
    Name of the GPO.
.PARAMETER All
    If true, gets all GPOs.
.PARAMETER Domain
    Active Directory domain to target.
.PARAMETER Server
    Domain Controller to target.
.PARAMETER Credential
    Alternate credentials to use.
.OUTPUTS
    The GPO.
#>
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='Guid')]
        [System.Guid]$Guid = [System.Guid]::Empty,

        [Parameter(Mandatory=$true, ParameterSetName='Name')]
        [String]$Name = $null,

        [Parameter(Mandatory=$true, ParameterSetName='All')]
        [Switch]$All = $false,

        [Parameter(Mandatory=$false)]
        [String]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($All) {
        $parameters.Add('All', $true)
    } elseif (![String]::IsNullOrEmpty($Name)) {
        $parameters.Add('Name', $Name)
    } else {
        $parameters.Add('Guid', $Guid)
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if (![String]::IsNullOrEmpty($Domain)) {
        $parameters.Add('Domain', $Domain)
    }
    $gpo = $null
    if ($Credential) {
        $properties = @(
            'DisplayName'
            'Id'
            'GpoStatus'
            'Description'
            'Path'
            'WmiFilter'
        )
        $ps = @{
            'ScriptBlock' = {
                Get-GPO @Using:parameters | Select-Object -Property $Using:properties
            }
            'Credential' = $Credential
            'ComputerName' = $env:COMPUTERNAME
            'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        }
        $gpo = Invoke-Command @ps
    } else {
        $gpo = Get-GPO @parameters
    }
    return $gpo
}

function Import-XGPO {
<#
.SYNOPSIS
    Imports the GPO settings from a backed-up GPO into a specified GPO.
.DESCRIPTION
    Imports the GPO settings from a backed-up GPO into a specified GPO.
.PARAMETER BackupId
    Guid of the backed up GPO.
.PARAMETER BackupGpoName
    Name of the backed up GPO.
.PARAMETER Path
    Specifies the path to the backup directory.
.PARAMETER TargetGuid
    GUID of the GPO into which this cmdlet imports the settings.
.PARAMETER TargetName
    Display name of the GPO into which the settings are to be imported.
.PARAMETER MigrationTable
    Path to a migration table file.
.PARAMETER CreateIfNeeded
    Creates a GPO from the backup if the target GPO does not exist.
.PARAMETER Domain
    Active Directory domain to target.
.PARAMETER Server
    DOmain Controller to target.
.PARAMETER Credential
    Alternate credentials to use.
.OUTPUTS
    Imported GPO.
#>
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='Guid')]
        [System.Guid]$BackupId = [System.Guid]::Empty,

        [Parameter(Mandatory=$true, ParameterSetName='Name')]
        [String]$BackupGpoName = $null,

        [Parameter(Mandatory=$true)]
        [String]$Path = $null,

        [Parameter(Mandatory=$false)]
        [System.Guid]$TargetGuid = [System.Guid]::Empty,

        [Parameter(Mandatory=$false)]
        [String]$TargetName = $null,

        [Parameter(Mandatory=$false)]
        [String]$MigrationTable = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$CreateIfNeeded = $false,

        [Parameter(Mandatory=$false)]
        [String]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'Path' = $Path
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($BackupId -ne [System.Guid]::Empty) {
        $parameters.Add('BackupId', $BackupId)
    } else {
        $parameters.Add('BackupGpoName', $BackupGpoName)
    }
    if ($TargetGuid -ne [System.Guid]::Empty) {
        $parameters.Add('TargetGuid', $TargetGuid)
    } elseif (![String]::IsNullOrEmpty($TargetName)) {
        $parameters.Add('TargetName', $TargetName)
    }
    if (![String]::IsNullOrEmpty($MigrationTable)) {
        $parameters.Add('MigrationTable', $MigrationTable)
    }
    if ($CreateIfNeeded) {
        $parameters.Add('CreateIfNeeded', $true)
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if (![String]::IsNullOrEmpty($Domain)) {
        $parameters.Add('Domain', $Domain)
    }
    $gpo = $null
    if ($Credential) {
        $properties = @(
            'DisplayName'
            'Id'
            'GpoStatus'
            'Description'
            'Path'
            'WmiFilter'
        )
        $ps = @{
            'ScriptBlock' = {
                Import-GPO @Using:parameters | Select-Object -Property $Using:properties
            }
            'Credential' = $Credential
            'ComputerName' = $env:COMPUTERNAME
            'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        }
        $gpo = Invoke-Command @ps
    } else {
        $gpo = Import-GPO @parameters
    }
    return $gpo
}

function Backup-XGPO {
<#
.SYNOPSIS
    Backs up one GPO or all the GPOs in a domain.
.DESCRIPTION
    Backs up one GPO or all the GPOs in a domain.
.PARAMETER Guid
    Guid of the GPO.
.PARAMETER Name
    Name of the GPO.
.PARAMETER All
    If true, backs up all GPOs.
.PARAMETER Path
    Specifies the path to the backup directory.
.PARAMETER Comment
    Specifies a comment for the backed-up GPO.
.PARAMETER Domain
    Active Directory domain to target.
.PARAMETER Server
    Domain Controller to target.
.PARAMETER Credential
    Alternate credentials to use.
.OUTPUTS
    The GPO.
#>
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='Guid')]
        [System.Guid]$Guid = [System.Guid]::Empty,

        [Parameter(Mandatory=$true, ParameterSetName='Name')]
        [String]$Name = $null,

        [Parameter(Mandatory=$true, ParameterSetName='All')]
        [Switch]$All = $false,

        [Parameter(Mandatory=$true)]
        [String]$Path = $false,

        [Parameter(Mandatory=$false)]
        [String]$Comment = $false,

        [Parameter(Mandatory=$false)]
        [String]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'Path' = $Path
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($All) {
        $parameters.Add('All', $true)
    } elseif ($Guid -ne [System.Guid]::Empty) {
        $parameters.Add('Guid', $Guid)
    } else {
        $parameters.Add('Name', $Name)
    }
    if (![String]::IsNullOrEmpty($Comment)) {
        $parameters.Add('Comment', $Comment)
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if (![String]::IsNullOrEmpty($Domain)) {
        $parameters.Add('Domain', $Domain)
    }
    $gpo = $null
    if ($Credential) {
        $properties = @(
            'DisplayName'
            'GpoId'
            'Id'
            'BackupDirectory'
            #'CreationTime'
            #'DomainName'
            #'Comment'
        )
        $ps = @{
            'ScriptBlock' = {
                Backup-GPO @Using:parameters | Select-Object -Property $Using:properties
            }
            'Credential' = $Credential
            'ComputerName' = $env:COMPUTERNAME
            'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        }
        $gpo = Invoke-Command @ps
    } else {
        $gpo = Backup-GPO @parameters
    }
    return $gpo
}

function Get-XGPInheritance {
<#
.SYNOPSIS
    Gets Group Policy inheritance information for a specified domain or OU.
.DESCRIPTION
    Gets Group Policy inheritance information for a specified domain or OU.
.PARAMETER Target
    OU or domain for which to retrieve the information.
.PARAMETER Domain
    Active Directory domain to target.
.PARAMETER Server
    Domain Controller to target.
.PARAMETER Credential
    Alternate credentials to use.
.OUTPUTS
    Group Policy inheritance information.
#>
    param(
        [Parameter(Mandatory=$true)]
        [String]$Target = $false,

        [Parameter(Mandatory=$false)]
        [String]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'Target' = $Target
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if (![String]::IsNullOrEmpty($Domain)) {
        $parameters.Add('Domain', $Domain)
    }
    $info = $null
    if ($Credential) {
        $properties = @(
            'Path'
            'GpoInheritanceBlocked'
            'GpoLinks'
        )
        $ps = @{
            'ScriptBlock' = {
                Get-GPInheritance @Using:parameters | Select-Object -Property $Using:properties
            }
            'Credential' = $Credential
            'ComputerName' = $env:COMPUTERNAME
            'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        }
        $info = Invoke-Command @ps
    } else {
        $info = Get-GPInheritance @parameters
    }
    return $info
}

function Get-XGPOReport {
<#
.SYNOPSIS
    Generates a report for a specified GPO or for all GPOs in a domain.
.DESCRIPTION
    Generates a report for a specified GPO or for all GPOs in a domain.
.PARAMETER Guid
    Guid of the GPO.
.PARAMETER Name
    Name of the GPO.
.PARAMETER All
    If true, gets all GPOs.
.PARAMETER ReportType
    Specifies the format of the report.
.PARAMETER Path
    Specifies the path to the report file.
.PARAMETER Domain
    Active Directory domain to target.
.PARAMETER Server
    Domain Controller to target.
.PARAMETER Credential
    Alternate credentials to use.
.OUTPUTS
    None.
#>
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='Guid')]
        [System.Guid]$Guid = [System.Guid]::Empty,

        [Parameter(Mandatory=$true, ParameterSetName='Name')]
        [String]$Name = $null,

        [Parameter(Mandatory=$true, ParameterSetName='All')]
        [Switch]$All = $false,

        [Parameter(Mandatory=$false)]
        [Microsoft.GroupPolicy.ReportType]$ReportType = [Microsoft.GroupPolicy.ReportType]::Html,

        [Parameter(Mandatory=$false)]
        [String]$Path = $null,

        [Parameter(Mandatory=$false)]
        [String]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'ReportType' = $ReportType
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($All) {
        $parameters.Add('All', $true)
    } elseif (![String]::IsNullOrEmpty($Name)) {
        $parameters.Add('Name', $Name)
    } else {
        $parameters.Add('Guid', $Guid)
    }
    if (![String]::IsNullOrEmpty($Path)) {
        $parameters.Add('Path', $Path)
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if (![String]::IsNullOrEmpty($Domain)) {
        $parameters.Add('Domain', $Domain)
    }
    if ($Credential) {
        $ps = @{
            'ScriptBlock' = { Get-GPOReport @Using:parameters }
            'Credential' = $Credential
            'ComputerName' = $env:COMPUTERNAME
            'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        }
        Invoke-Command @ps
    } else {
        Get-GPOReport @parameters
    }
}

function New-XGPLink {
<#
.SYNOPSIS
    Links a GPO.
.DESCRIPTION
    Links a GPO.
.PARAMETER Guid
    Guid of the GPO.
.PARAMETER Name
    Name of the GPO.
.PARAMETER Target
    Container to which to link the GPO.
.PARAMETER LinkEnabled
    True if the GPO's Link must be enabled.
.PARAMETER Order
    Link Order.
.PARAMETER Enforced
    True if the GPO's Link must be enforced.
.PARAMETER Domain
    Active Directory domain to target.
.PARAMETER Server
    Domain Controller to target.
.PARAMETER Credential
    Alternate credentials to use.
.OUTPUTS
    Link.
#>
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='Guid')]
        [System.Guid]$Guid = [System.Guid]::Empty,

        [Parameter(Mandatory=$true, ParameterSetName='Name')]
        [String]$Name = $null,

        [Parameter(Mandatory=$true)]
        [String]$Target = $null,

        [Parameter(Mandatory=$false)]
        [Microsoft.GroupPolicy.EnableLink]$LinkEnabled = [Microsoft.GroupPolicy.EnableLink]::No,

        [Parameter(Mandatory=$false)]
        [Int32]$Order = 0,

        [Parameter(Mandatory=$false)]
        [Microsoft.GroupPolicy.EnforceLink]$Enforced = [Microsoft.GroupPolicy.EnforceLink]::No,

        [Parameter(Mandatory=$false)]
        [String]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'Target' = $Target
        'LinkEnabled' = $LinkEnabled
        'Enforced' = $Enforced
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($Guid -ne [System.Guid]::Empty) {
        $parameters.Add('Guid', $Guid)
    } else {
        $parameters.Add('Name', $Name)
    }
    if ($Order -gt 0) {
        $parameters.Add('Order', $Order)
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if (![String]::IsNullOrEmpty($Domain)) {
        $parameters.Add('Domain', $Domain)
    }
    $link = $null
    if ($Credential) {
        $properties = @(
            'GpoId'
            'DisplayName'
            'Enabled'
            'Enforced'
            'Target'
            'Order'
        )
        $ps = @{
            'ScriptBlock' = {
                New-GPLink @Using:parameters | Select-Object -Property $Using:properties
            }
            'Credential' = $Credential
            'ComputerName' = $env:COMPUTERNAME
            'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        }
        $link = Invoke-Command @ps
    } else {
        $link = New-GPLink @parameters
    }
    return $link
}

function Set-XGPLink {
<#
.SYNOPSIS
    Sets the properties of the specified GPO link.
.DESCRIPTION
    Sets the properties of the specified GPO link.
.PARAMETER Guid
    Guid of the GPO.
.PARAMETER Name
    Name of the GPO.
.PARAMETER Target
    Container to which to link the GPO.
.PARAMETER LinkEnabled
    True if the GPO's Link must be enabled.
.PARAMETER Order
    Link Order.
.PARAMETER Enforced
    True if the GPO's Link must be enforced.
.PARAMETER Domain
    Active Directory domain to target.
.PARAMETER Server
    Domain Controller to target.
.PARAMETER Credential
    Alternate credentials to use.
.OUTPUTS
    Link.
#>
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='Guid')]
        [System.Guid]$Guid = [System.Guid]::Empty,

        [Parameter(Mandatory=$true, ParameterSetName='Name')]
        [String]$Name = $null,

        [Parameter(Mandatory=$true)]
        [String]$Target = $null,

        [Parameter(Mandatory=$false)]
        [Microsoft.GroupPolicy.EnableLink]$LinkEnabled = [Microsoft.GroupPolicy.EnableLink]::No,

        [Parameter(Mandatory=$false)]
        [Int32]$Order = 0,

        [Parameter(Mandatory=$false)]
        [Microsoft.GroupPolicy.EnforceLink]$Enforced = [Microsoft.GroupPolicy.EnforceLink]::No,

        [Parameter(Mandatory=$false)]
        [String]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'Target' = $Target
        'LinkEnabled' = $LinkEnabled
        'Enforced' = $Enforced
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($Guid -ne [System.Guid]::Empty) {
        $parameters.Add('Guid', $Guid)
    } else {
        $parameters.Add('Name', $Name)
    }
    if ($Order -gt 0) {
        $parameters.Add('Order', $Order)
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if (![String]::IsNullOrEmpty($Domain)) {
        $parameters.Add('Domain', $Domain)
    }
    $link = $null
    if ($Credential) {
        $properties = @(
            'GpoId'
            'DisplayName'
            'Enabled'
            'Enforced'
            'Target'
            'Order'
        )
        $ps = @{
            'ScriptBlock' = {
                Set-GPLink @Using:parameters | Select-Object -Property $Using:properties
            }
            'Credential' = $Credential
            'ComputerName' = $env:COMPUTERNAME
            'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        }
        $link = Invoke-Command @ps
    } else {
        $link = Set-GPLink @parameters
    }
    return $link
}

function Remove-XGPLink {
<#
.SYNOPSIS
    Removes a GPO link.
.DESCRIPTION
    Removes a GPO link.
.PARAMETER Guid
    Guid of the GPO.
.PARAMETER Name
    Name of the GPO.
.PARAMETER Target
    Container from which to link is removed.
.PARAMETER Domain
    Active Directory domain to target.
.PARAMETER Server
    Domain Controller to target.
.PARAMETER Credential
    Alternate credentials to use.
.OUTPUTS
    Link.
#>
    [CmdletBinding(DefaultParameterSetName='Name')]
    param(
        [Parameter(Mandatory=$true, ParameterSetName='Guid')]
        [System.Guid]$Guid = [System.Guid]::Empty,

        [Parameter(Mandatory=$true, ParameterSetName='Name')]
        [String]$Name = $null,

        [Parameter(Mandatory=$true)]
        [String]$Target = $null,

        [Parameter(Mandatory=$false)]
        [String]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'Target' = $Target
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($Guid -ne [System.Guid]::Empty) {
        $parameters.Add('Guid', $Guid)
    } else {
        $parameters.Add('Name', $Name)
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if (![String]::IsNullOrEmpty($Domain)) {
        $parameters.Add('Domain', $Domain)
    }
    $link = $null
    if ($Credential) {
        $properties = @(
            'GpoId'
            'DisplayName'
            'Enabled'
            'Enforced'
            'Target'
            'Order'
        )
        $ps = @{
            'ScriptBlock' = {
                Remove-GPLink @Using:parameters | Select-Object -Property $Using:properties
            }
            'Credential' = $Credential
            'ComputerName' = $env:COMPUTERNAME
            'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
        }
        $link = Invoke-Command @ps
    } else {
        $link = Remove-GPLink @parameters
    }
    return $link
}

function Get-PathFromDn {
<#
.SYNOPSIS
    Gets a Path from a DistinguishedName.
.DESCRIPTION
    Gets a Path from a DistinguishedName.
.PARAMETER Dn
    DistinguishedName.
.OUTPUTS
    Path of the object.
#>
    param (
        [parameter(mandatory=$false)]
        [String]$Dn = $null
    )
    if ([String]::IsNullOrEmpty($Dn) -or ($Dn -inotmatch '^(CN=|OU=).*')) {
        return $null
    }
    $Dn = $Dn.Substring(3)
    $chars = $Dn.ToCharArray()
    for ($i = 0; $i -lt $chars.Count; $i++) {
        if ($chars[$i] -eq ',') {
            $j = $i - 1
            $count = 0
            while (($j -ge 0) -and ($chars[$j] -eq '\')) {
                $j--
                $count++
            }
            if (($count % 2) -eq 0) {
                return $Dn.Substring($i + 1)
            }
        }
    }
    return $null
}

function Get-AnonymizedSid {
<#
.SYNOPSIS
    Anonymizes a SID.
.DESCRIPTION
    Anonymizes a SID.
.PARAMETER Sid
    SID.
.PARAMETER Domain
    Active Directory domain object.
.OUTPUTS
    Anonymized SID.
#>
    param (
        [parameter(mandatory=$true)]
        [String]$Sid = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null
    )
    $patternDomSid = '^(S-1-5-21-\d+-\d+-\d+)-([1-9]\d*)$'
    if ($Sid -imatch $patternDomSid) {
        $forestLevel = @('498', '518', '519', '527')
        if ($Matches[1] -ieq $Domain.DomainSID.ToString()) {
            $Sid = "[DomainSID]-$($Matches[2])"
            if ($Matches[2] -iin $forestLevel) {
                $Sid = "[RootDomainSID]-$($Matches[2])"
            }
        } elseif ($Matches[1] -ieq $Domain.RootDomainSID.ToString()) {
            $Sid = "[RootDomainSID]-$($Matches[2])"
        } else {
            $Sid = "[RemoteDomainSID]-$($Matches[2])"
        }
    }
    return $Sid
}

function Get-SystemContainer {
<#
.SYNOPSIS
    Gets a system container.
.DESCRIPTION
    Gets a system container.
.PARAMETER ContainerId
    Identifier of the system container.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to query.
.PARAMETER Credential
    Credential to use to query the Domain Controller.
.OUTPUTS
    DistinguishedName of the System containers or null.
#>
    param (
        [Parameter(Mandatory=$false)]
        [String]$ContainerId = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    $parameters = @{
        'Identity' = $Domain.DistinguishedName
        'Properties' = @('wellKnownObjects', 'otherWellKnownObjects')
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if ($Credential) {
        $parameters.Add('Credential', $Credential)
    }
    $h = @{}
    try {
        $obj = Get-ADObject @parameters
        $list = @() + $obj.wellKnownObjects.Split(';')
        $list += $obj.otherWellKnownObjects.Split(';')
        foreach ($entry in $list) {
            if ($entry -imatch '^B:\d+:([^:]+):(.+)$') {
                $container = $Matches[1]
                $dn = $Matches[2]
                if ($container -ieq $ContainerId) {
                    return $dn
                }
                foreach ($c in $Script:SystemContainersId.Keys) {
                    $id = $Script:SystemContainersId[$c]
                    if (($id -ieq $container) -and !$h.ContainsKey($c)) {
                        $h.Add($c, $dn)
                    }
                }
            }
        }
    } catch {
        Add-Log -Log "Error while listing System Containers. $_" -Type Error
        $h = $null
    }
    return $h
}

#endregion

#region Xml Helpers

function Test-XmlSchema {
<#
.SYNOPSIS
    Validates the schema of an Xml file.
.DESCRIPTION
    Validates the schema of an Xml file.
.PARAMETER File
    Path to Xml file to be validated.
.OUTPUTS
    True on success, false otherwise.
#>
    param (
        [parameter(mandatory=$true)]
        [String]$File = $null
    )
    Add-Log -Log 'Validating Xml schema'
    if ([String]::IsNullOrEmpty($File)) {
        Add-Log -Log 'File can not be empty' -Type Error
        return $false
    }
    if (!(Test-Path -Path $File)) {
        Add-Log -Log "File '$File' not found" -Type Error
        return $false
    }
    $success = $true
    try {
        $prop = Get-ItemProperty -Path $File -ErrorAction Stop
        if ($prop.FullName.StartsWith('\\')) {
            Copy-Item -Path $File -Destination $Script:TempFolder -ErrorAction Stop
            $File = Join-Path -Path $Script:TempFolder -ChildPath $prop.Name
        } else {
            $File = $prop.FullName
        }
    } catch {
        Add-Log -Log "Error while resolving path. $_" -Type Error
        return $false
    }
    $settings = New-Object -TypeName System.Xml.XmlReaderSettings
    $settings.ValidationType = [System.Xml.ValidationType]::Schema
    $settings.ValidationFlags =
        [System.Xml.Schema.XmlSchemaValidationFlags]::ProcessInlineSchema -bor
        [System.Xml.Schema.XmlSchemaValidationFlags]::ProcessIdentityConstraints -bor
        [System.Xml.Schema.XmlSchemaValidationFlags]::ProcessSchemaLocation -bor
        [System.Xml.Schema.XmlSchemaValidationFlags]::ReportValidationWarnings
    $xmlReader = $null
    try {
        $strReader = New-Object -TypeName System.IO.StringReader -ArgumentList $Script:XmlSchema
        $xmlReader = [System.Xml.XmlReader]::Create($strReader, $settings)
        $settings.Schemas.Add($null, $xmlReader) | Out-Null
        $settings.add_ValidationEventHandler({
            throw "$($_.Message) Line $($_.Exception.LineNumber), Col $($_.Exception.LinePosition))"
        })
    } catch {
        Add-Log -Log "Invalid Xml Schema. $_" -Type Error
        $success = $false
    } finally {
        if ($xmlReader) {
            $xmlReader.Close()
            $xmlReader = $null
        }
    }
    if ($success) {
        try {
            $xmlReader = [System.Xml.XmlReader]::Create($File, $settings)
            while ($xmlReader.Read()) {
            }
        } catch {
            Add-Log -Log "Error while validating Xml file. $_" -Type Error
            $success = $false
        } finally {
            if ($xmlReader) {
                $xmlReader.Close()
            }
        }
    }
    return $success
}

function Get-XmlValue {
<#
.SYNOPSIS
    Gets Xml value from an Xml Element.
.DESCRIPTION
    Gets Xml value from an Xml Element.
.PARAMETER Xml
    Xml Element.
.PARAMETER XmlPath
    Xml path.
.PARAMETER DefaultValue
    Default value.
.PARAMETER DisplayError
    If true, error message is shown.
.OUTPUTS
    Value of the Xml element or null.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$Xml = $null,

        [Parameter(Mandatory=$true)]
        [String]$XmlPath = $null,

        [Parameter(Mandatory=$false)]
        [Object]$DefaultValue = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$DisplayError = $false
    )
    if (!$Xml) {
        return $null
    }
    $data = $DefaultValue
    try {
        $data = $Xml.$XmlPath
    } catch {
        if ($DisplayError) {
            Add-Log -Log "Xml Path '$($Xml.Name).$($XmlPath)' not found, default value used instead" -Type Warning
        }
        return $DefaultValue
    }
    return $data
}

function Read-XmlSettings {
<#
.SYNOPSIS
    Reads Settings Xml file.
.DESCRIPTION
    Reads Settings Xml file.
.PARAMETER Path
    Path to the Xml file.
.OUTPUTS
    Xml object or null.
#>
    param (
        [Parameter(Mandatory=$true)]
        [String]$Path = $null
    )
    if ([String]::IsNullOrEmpty($Path) -or
        !(Test-Path -Path $Path -IsValid) -or
        !(Test-Path -Path $Path)) {
        Add-Log -Log 'Path parameter must be a valid path' -Type Error
        return $null
    }
    try {
        $Path = (Resolve-Path -Path $Path -ErrorAction Stop).Path
    } catch {
        Add-Log -Log "Unable to resolve path. $_" -Type Error
        return $null
    }
    Add-Log -Log "Xml file: '$Path'"
    if (!(Test-XmlSchema -File $Path)) {
        return $null
    }
    Add-Log -Log 'Parsing Xml file'
    $xml = $null
    try {
        [Xml]$xml = Get-Content -Path $Path -ErrorAction Stop
    } catch {
        Add-Log -Log "Can not read file '$Path'. $_" -Type Error
        return $null
    }
    return $xml
}

function Create-SettingsXmlNode {
<#
.SYNOPSIS
    Creates an Xml node.
.DESCRIPTION
    Creates an Xml node.
.PARAMETER XmlDoc
    Xml Document.
.PARAMETER NodeName
    Name of the Xml node to be created.
.PARAMETER Data
    Object from which to extract data.
.PARAMETER DefaultPermissions
    Default permissions on AD object.
.PARAMETER Domain
    Active Directory Domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    False if an error occurrs, Xml Element or null otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$XmlDoc = $null,

        [Parameter(Mandatory=$true)]
        [String]$NodeName = $null,

        [Parameter(Mandatory=$false)]
        [Object]$Data = $null,

        [Parameter(Mandatory=$false)]
        [System.Security.AccessControl.AuthorizationRule[]]$DefaultPermissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    if (!$XmlDoc -or [String]::IsNullOrEmpty($NodeName)) {
        Add-Log -Log 'A valid Xml Document and Xml node must be provided' -Type Error
        return $false
    }
    if (!$Data) {
        #Add-Log -Log "No Data for node '$($NodeName)'" -Type Warning
    }
    $commonParams = @{
        'Domain' = $Domain
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    $x = $xmlDoc.CreateNode('element', $NodeName, $null)
    $v = ''
    switch ($NodeName) {
        $Script:DefSystemContainer {
            $v = ($data -join ' ')
            break
        }
        $Script:DefDistinguishedName {
            $v = "$($Data)".Trim()
            if ($v -imatch "(.*?)\s*,?\s*$($domain.DistinguishedName)$") {
                $v = $Matches[1]
            }
            break
        }
        $Script:DefGpoInheritanceBlocked {
            $v = ($Data -eq 1).ToString()
            break
        }
        $Script:DefMembers {
            if (!$Data) {
                return $null
            }
            foreach ($member in $Data) {
                if (!$member) {
                    continue
                }
                $xMember = $xmlDoc.CreateNode('element', $Script:DefMember, $null)
                $xMember.InnerText = $member.SamAccountName
                $xMember.SetAttribute($Script:DefIsWellKnown, 'false') | Out-Null
                $sid = $member.SID
                if ($sid -and (Test-WellKnownSid -Sid $sid)) {
                    $rid = Get-AnonymizedSid -Sid $sid -Domain $Domain
                    $xMember.SetAttribute($Script:DefIsWellKnown, 'true') | Out-Null
                    $xMember.InnerText = $rid
                }
                $x.AppendChild($xMember) | Out-Null
            }
            break
        }
        $Script:DefLinks {
            foreach ($link in $Data) {
                $xLink = $xmlDoc.CreateNode('element', $Script:DefLink, $null)

                $n = $xmlDoc.CreateNode('element', $Script:DefEnabled, $null)
                $n.InnerText = $link.Enabled.ToString().ToLower()
                $xLink.AppendChild($n) | Out-Null

                $n = $xmlDoc.CreateNode('element', $Script:DefEnforced, $null)
                $n.InnerText = $link.Enforced.ToString().ToLower()
                $xLink.AppendChild($n) | Out-Null

                $n = $xmlDoc.CreateNode('element', $Script:DefOrder, $null)
                $n.InnerText = $link.Order
                $xLink.AppendChild($n) | Out-Null

                $n = $xmlDoc.CreateNode('element', $Script:DefContainer, $null)
                $dn = $link.Target.Trim()
                if ($dn -imatch "(.*?)\s*,?\s*$($domain.DistinguishedName)$") {
                    $dn = $Matches[1]
                }
                $n.InnerText = $dn
                $xLink.AppendChild($n) | Out-Null

                $x.AppendChild($xLink) | Out-Null
            }
            break
        }
        $Script:DefPermissions {
            if (!$Data) {
                return $null
            }
            foreach ($acl in $Data.Access) {
                if (!$acl -or $acl.IsInherited) {
                    continue
                }
                if ((Test-ProtectedFromDeletion -Rules $acl)) {
                    continue
                }
                if ((Test-AclInList -Challengers $acl `
                                    -References $DefaultPermissions `
                                    -Server $Server `
                                    -Credential $Credential `
                                    -Domain $Domain)) {
                    continue
                }
                $xPerm = $xmlDoc.CreateNode('element', $Script:DefPermission, $null)
                $n = $xmlDoc.CreateNode('element', $Script:DefRights, $null)
                $n.InnerText = $acl.ActiveDirectoryRights.ToString().Replace(', ', ' ')
                $xPerm.AppendChild($n) | Out-Null
                $n = $xmlDoc.CreateNode('element', $Script:DefInheritance, $null)
                $n.InnerText = $acl.InheritanceType.ToString()
                $xPerm.AppendChild($n) | Out-Null
                if ($acl.ObjectType -ne [System.Guid]::Empty) {
                    $n = $xmlDoc.CreateNode('element', $Script:DefObject, $null)
                    $n.InnerText = $acl.ObjectType.ToString().ToLower()
                    $xPerm.AppendChild($n) | Out-Null
                }
                if ($acl.InheritedObjectType -ne [System.Guid]::Empty) {
                    $n = $xmlDoc.CreateNode('element', $Script:DefInheritedObject, $null)
                    $n.InnerText = $acl.InheritedObjectType.ToString().ToLower()
                    $xPerm.AppendChild($n) | Out-Null
                }
                $n = $xmlDoc.CreateNode('element', $Script:DefAccessControl, $null)
                $n.InnerText = $acl.AccessControlType.ToString()
                $xPerm.AppendChild($n) | Out-Null
                $n = $xmlDoc.CreateNode('element', $Script:DefIdentityReference, $null)
                $n.InnerText = $acl.IdentityReference.Value.Split('\') | Select-Object -Last 1
                $n.SetAttribute($Script:DefIsWellKnown, 'false') | Out-Null
                $sid = Convert-StringToSid -String $acl.IdentityReference.Value @commonParams
                if ($sid -and (Test-WellKnownSid -Sid $sid)) {
                    $rid = Get-AnonymizedSid -Sid $sid -Domain $Domain
                    $n.SetAttribute($Script:DefIsWellKnown, 'true') | Out-Null
                    $n.InnerText = $rid
                }
                $xPerm.AppendChild($n) | Out-Null
                $x.AppendChild($xPerm) | Out-Null
            }
            break
        }
        $Script:DefUserPrincipalName {
            $v = ("$($Data)" -isplit '@')[0]
            break
        }
        default {
            $v = "$($Data)"
            break
        }
    }
    if (!$x.HasChildNodes -and [String]::IsNullOrEmpty($v)) {
        return $null
    }
    if ($v -iin @('True', 'False')) {
        $v = $v.ToLower()
    }
    if (!$x.HasChildNodes) {
        $x.InnerText = $v
    }
    return $x
}

#endregion

#region Restore

function Restore-OrganizationalUnits {
<#
.SYNOPSIS
    Restores Organizational Unit objects.
.DESCRIPTION
    Restores Organizational Unit objects.
.PARAMETER Data
    Data to restore.
.PARAMETER RedirectContainer
    If true, system container is redirected.
.PARAMETER Permissions
    Variable listing permissions (out).
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$false)]
        [System.Xml.XmlElement]$Data = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$RedirectContainer = $false,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$Permissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    Add-Log -Log 'Restoring Organizational Unit objects' -Type Title2
    if (!$Data) {
        Add-Log -Log 'No Organizational Unit object to restore'
        return $true
    }
    if (!$Domain) {
        Add-Log -Log 'Active Directory object cannot be null' -Type Error
        return $false
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    if (!$Permissions) {
        $Permissions = @{}
    }
    foreach ($obj in $Data.ChildNodes) {
        $dn = Get-XmlValue -Xml $obj -XmlPath $Script:DefDistinguishedName
        if ([String]::IsNullOrEmpty($dn)) {
            Add-Log -Log 'DistinguishedName cannot be empty (skipped)' -Type Warning
            continue
        }
        $dn += ",$($Domain.DistinguishedName)"
        $dnLc = $dn.ToLower()
        if (![String]::IsNullOrEmpty($Scope) -and !($dnLc.EndsWith($Scope.ToLower()))) {
            Add-Log -Log "Object '$dn' is not in the scope '$Scope' (skipped)" -Type Warning
            continue
        }
        $name = Get-XmlValue -Xml $obj -XmlPath $Script:DefName
        if (($name.Length + 4) -ge $dn.Length) {
            Add-Log -Log "Invalid object's name '$name' (skipped)" -Type Warning
            continue
        }
        $path = Get-PathFromDn -Dn $dn
        if ([String]::IsNullOrEmpty($path)) {
            Add-Log -Log "Invalid path for object '$dn' (skipped)" -Type Warning
            continue
        }
        $systemContainer = Get-XmlValue -Xml $obj -XmlPath $Script:DefSystemContainer
        $description = Get-XmlValue -Xml $obj -XmlPath $Script:DefDescription
        $protected = Get-XmlValue -Xml $obj -XmlPath $Script:DefProtectedFromDel -DefaultValue 'true'
        $gpoInheritanceBlocked = Get-XmlValue -Xml $obj -XmlPath $Script:DefGpoInheritanceBlocked -DefaultValue 'false'
        $gpoInheritanceBlocked = ($gpoInheritanceBlocked -ieq 'true')
        $perms = Get-XmlValue -Xml $obj -XmlPath $Script:DefPermissions
        if ($perms) {
            if ($Permissions.ContainsKey($dnLc)) {
                $Permissions[$dnLc] += $perms
            } else {
                $Permissions.Add($dnLc, $perms)
            }
        }
        $ps = @{
            'Description' = $description
            'ProtectedFromAccidentalDeletion' = ($protected -ieq 'true')
        }
        $clear = @()
        $parameters = @{}
        foreach ($key in $ps.Keys) {
            if ($ps[$key] -eq $null) {
                continue
            }
            if (($ps[$key] -is [String]) -and ($ps[$key] -eq '')) {
                $clear += $key
            } else {
                $parameters.Add($key, $ps[$key])
            }
        }
        $ou = $null
        try {
            if ($dnLc.StartsWith('cn=')) {
                $ou = Get-ADObject -Identity $dn -Properties 'DistinguishedName', 'gPLink' @commonParams
            } else {
                $ou = Get-ADOrganizationalUnit -Identity $dn -Properties 'DistinguishedName', 'gPLink' @commonParams
            }
        } catch {
            if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                Add-Log -Log "Error while getting object '$dn' information. $_" -Type Error
                return $false
            }
        }
        $parameters += $commonParams
        $replace = @{
            'gPOptions' = 0
        }
        if ($gpoInheritanceBlocked) {
            $replace['gPOptions'] = 1
            if (!$ou -or ($ou -and [String]::IsNullOrEmpty($ou.gPLink))) {
                $replace.Add('gPLink', ' ')
            }
        }
        try {
            if ($ou) {
                if ($Force -and !$dnLc.StartsWith('cn=')) {
                    Add-Log -Log "Updating object '$ou' (name=$name)"
                    if ($clear.Count -gt 0) {
                        $parameters.Add('Clear', $clear)
                    }
                    Set-ADOrganizationalUnit -Identity $dn -Replace $replace @parameters
                } else {
                    Add-Log -Log "Object '$dn' already exists"
                }
            } else {
                Add-Log -Log "Creating object '$dn'"
                New-ADOrganizationalUnit -Name $name `
                                         -Path $path `
                                         -OtherAttributes $replace `
                                         @parameters
            }
        } catch {
            Add-Log -Log "Error while restoring object '$dn'. $_" -Type Error
            return $false
        }
        if ($RedirectContainer -and ![String]::IsNullOrEmpty($systemContainer)) {
            if (!$Script:SystemContainersId.ContainsKey($systemContainer)) {
                Add-Log -Log "Unknown system container '$systemContainer'" -Type Warning
                continue
            }
            $containerId = $Script:SystemContainersId[$systemContainer]
            $currentContainer = Get-SystemContainer -ContainerId $containerId `
                                                    -Domain $Domain `
                                                    -Server $Server `
                                                    -Credential $Credential
            if ([String]::IsNullOrEmpty($currentContainer)) {
                Add-Log -Log "System container '$systemContainer' not found" -Type Warning
                continue
            }
            try {
                $attr = 'wellKnownObjects'
                if ($systemContainer -iin @('ManagedServiceAccounts')) {
                    $attr = 'otherWellKnownObjects'
                }
                if ($Force -or ($dn -ine $currentContainer)) {
                    $old = "B:$($containerId.Length):$($containerId):$currentContainer"
                    $new = "B:$($containerId.Length):$($containerId):$dn"
                    Set-ADObject -Identity $Domain.DistinguishedName `
                                 -Add @{ $attr = $new } `
                                 -Remove @{ $attr = $old } `
                                 @commonParams
                    Add-Log -Log "Redirected system container '$($systemContainer)' to '$($dn)'"
                } else {
                    Add-Log -Log "System container '$($systemContainer)' already redirected to '$($dn)'"
                }
            } catch {
                Add-Log -Log "Unable to redirect system container '$($systemContainer)' to '$($dn)'. $_" -Type Error
                return $false
            }
        }
    }
    return $true
}

function Restore-Users {
<#
.SYNOPSIS
    Restores User objects.
.DESCRIPTION
    Restores User objects.
.PARAMETER Data
    Data to restore.
.PARAMETER Permissions
    Variable listing permissions (out).
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$false)]
        [System.Xml.XmlElement]$Data = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$Permissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    Add-Log -Log 'Restoring User Objects' -Type Title2
    if (!$Data) {
        Add-Log -Log 'No User object to restore'
        return $true
    }
    if (!$Domain) {
        Add-Log -Log 'Active Directory object cannot be null' -Type Error
        return $false
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    if (!$Permissions) {
        $Permissions = @{}
    }
    foreach ($obj in $Data.ChildNodes) {
        $dn = Get-XmlValue -Xml $obj -XmlPath $Script:DefDistinguishedName
        if ([String]::IsNullOrEmpty($dn)) {
            Add-Log -Log 'DistinguishedName cannot be empty (skipped)' -Type Warning
            continue
        }
        $dn += ",$($Domain.DistinguishedName)"
        $dnLc = $dn.ToLower()
        if (![String]::IsNullOrEmpty($Scope) -and !($dnLc.EndsWith($Scope.ToLower()))) {
            Add-Log -Log "Object '$dn' is not in the scope '$Scope' (skipped)" -Type Warning
            continue
        }
        $name = Get-XmlValue -Xml $obj -XmlPath $Script:DefName
        if (($name.Length + 4) -ge $dn.Length) {
            Add-Log -Log "Invalid object's name '$name' (skipped)" -Type Warning
            continue
        }
        $path = Get-PathFromDn -Dn $dn
        if ([String]::IsNullOrEmpty($path)) {
            Add-Log -Log "Invalid path for object '$dn' (skipped)" -Type Warning
            continue
        }
        $accountExpirationDate = Get-XmlValue -Xml $obj -XmlPath $Script:DefAccountExpirationDate
        try {
            if ($accountExpirationDate -ne $null) {
                if ($accountExpirationDate -match '^\d+$') {
                    $accountExpirationDate = [DateTime]::FromFileTime($accountExpirationDate)
                } else {
                    $accountExpirationDate = [DateTime]$accountExpirationDate
                }
            }
        } catch {
            $accountExpirationDate = 0
        }
        $accountNotDelegated = Get-XmlValue -Xml $obj -XmlPath $Script:DefAccountNotDelegated -DefaultValue 'false'
        $accountPassword = Get-ComplexPassword
        $allowRevPwdEncryption = Get-XmlValue -Xml $obj -XmlPath $Script:DefAllowRevPwdEncryption -DefaultValue 'false'
        $cannotChangePassword = Get-XmlValue -Xml $obj -XmlPath $Script:DefCannotChangePassword -DefaultValue 'false'
        $changePasswordAtLogon = Get-XmlValue -Xml $obj -XmlPath $Script:DefChangePasswordAtLogon -DefaultValue 'true'
        $description = Get-XmlValue -Xml $obj -XmlPath $Script:DefDescription
        $displayName = Get-XmlValue -Xml $obj -XmlPath $Script:DefDisplayName
        $enabled = Get-XmlValue -Xml $obj -XmlPath $Script:DefEnabled -DefaultValue 'true'
        $givenName = Get-XmlValue -Xml $obj -XmlPath $Script:DefGivenName
        $initials = Get-XmlValue -Xml $obj -XmlPath $Script:DefInitials
        $kerberosEncryptionType = Get-XmlValue -Xml $obj -XmlPath $Script:DefKerberosEncryptionType
        $passwordNeverExpires = Get-XmlValue -Xml $obj -XmlPath $Script:DefPasswordNeverExpires -DefaultValue 'false'
        $passwordNotRequired = Get-XmlValue -Xml $obj -XmlPath $Script:DefPasswordNotRequired -DefaultValue 'false'
        $samAccountName = Get-XmlValue -Xml $obj -XmlPath $Script:DefSamAccountName
        $scril = Get-XmlValue -Xml $obj -XmlPath $Script:DefSmartcardLogonRequired -DefaultValue 'false'
        $surname = Get-XmlValue -Xml $obj -XmlPath $Script:DefSurname
        $trustedForDelegation = Get-XmlValue -Xml $obj -XmlPath $Script:DefTrustedForDelegation -DefaultValue 'false'
        $userPrincipalName = Get-XmlValue -Xml $obj -XmlPath $Script:DefUserPrincipalName
        if ($passwordNeverExpires -ieq 'true') {
            $changePasswordAtLogon = 'false'
        }
        $perms = Get-XmlValue -Xml $obj -XmlPath $Script:DefPermissions
        if ($perms) {
            if ($Permissions.ContainsKey($dnLc)) {
                $Permissions[$dnLc] += $perms
            } else {
                $Permissions.Add($dnLc, $perms)
            }
        }
        $ps = @{
            'AccountNotDelegated' = ($accountNotDelegated -ieq 'true')
            'AllowReversiblePasswordEncryption' = ($allowRevPwdEncryption -ieq 'true')
            'CannotChangePassword' = ($cannotChangePassword -ieq 'true')
            'ChangePasswordAtLogon' = ($changePasswordAtLogon -ieq 'true')
            'Description' = $description
            'DisplayName' = $displayName
            'Enabled' = ($enabled -ieq 'true')
            'GivenName' = $givenName
            'Initials' = $initials
            #'KerberosEncryptionType' = $kerberosEncryptionType
            'PasswordNeverExpires' = ($passwordNeverExpires -ieq 'true')
            'PasswordNotRequired' = ($passwordNotRequired -ieq 'true')
            'SamAccountName' = $samAccountName
            'SmartcardLogonRequired' = ($scril -ieq 'true')
            'Surname' = $surname
            'TrustedForDelegation' = ($trustedForDelegation -ieq 'true')
            'UserPrincipalName' = $userPrincipalName
        }
        $clear = @()
        $parameters = @{}
        foreach ($key in $ps.Keys) {
            if ($ps[$key] -eq $null) {
                continue
            }
            if (($ps[$key] -is [String]) -and ($ps[$key] -eq '')) {
                switch ($key) {
                    'Surname' {
                        $clear += 'sn'
                        break
                    }
                    default {
                        $clear += $key
                        break
                    }
                }
            } else {
                $parameters.Add($key, $ps[$key])
            }
        }
        $user = $null
        try {
            $user = Get-ADUser -Identity $dn -Properties 'samAccountName', 'SID' @commonParams
        } catch {
            if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                Add-Log -Log "Error while getting object '$dn' information. $_" -Type Error
                return $false
            }
        }
        if ($user) {
            if (!(Test-WellKnownSid -Sid $user.SID)) {
                if ($accountExpirationDate -eq $null) {
                    $accountExpirationDate = 0
                }
                $parameters.Add('Replace', @{ 'accountExpires' = $accountExpirationDate })
            } else {
                if ($accountExpirationDate -ne $null) {
                    Add-Log -Log "Expiration date ignored because '$dn' is a built-in account" -Type Warning
                }
            }
        }
        $parameters += $commonParams
        try {
            if ($user) {
                if ($Force) {
                    Add-Log -Log "Updating object '$dn'"
                    if ($clear.Count -gt 0) {
                        $parameters.Add('Clear', $clear)
                    }
                    Set-ADUser -Identity $dn @parameters
                } else {
                    Add-Log -Log "Object '$dn' already exists"
                }
            } else {
                New-ADUser -Name $name -Path $path -AccountPassword $accountPassword @parameters
                Add-Log -Log "Created object '$dn'"
            }
        } catch {
            Add-Log -Log "Error while restoring object '$dn'. $_" -Type Error
            return $false
        }
    }
    return $true
}

function Restore-Computers {
<#
.SYNOPSIS
    Restores Computer objects.
.DESCRIPTION
    Restores Computer objects.
.PARAMETER Data
    Data to restore.
.PARAMETER Permissions
    Variable listing permissions (out).
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$false)]
        [System.Xml.XmlElement]$Data = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$Permissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    Add-Log -Log 'Restoring Computer Objects' -Type Title2
    if (!$Data) {
        Add-Log -Log 'No Computer object to restore'
        return $true
    }
    if (!$Domain) {
        Add-Log -Log 'Active Directory object cannot be null' -Type Error
        return $false
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    if (!$Permissions) {
        $Permissions = @{}
    }
    foreach ($obj in $Data.ChildNodes) {
        $dn = Get-XmlValue -Xml $obj -XmlPath $Script:DefDistinguishedName
        if ([String]::IsNullOrEmpty($dn)) {
            Add-Log -Log 'DistinguishedName cannot be empty (skipped)' -Type Warning
            continue
        }
        $dn += ",$($Domain.DistinguishedName)"
        $dnLc = $dn.ToLower()
        if (![String]::IsNullOrEmpty($Scope) -and !($dnLc.EndsWith($Scope.ToLower()))) {
            Add-Log -Log "Object '$dn' is not in the scope '$Scope' (skipped)" -Type Warning
            continue
        }
        $name = Get-XmlValue -Xml $obj -XmlPath $Script:DefName
        if (($name.Length + 4) -ge $dn.Length) {
            Add-Log -Log "Invalid object's name '$name' (skipped)" -Type Warning
            continue
        }
        $path = Get-PathFromDn -Dn $dn
        if ([String]::IsNullOrEmpty($path)) {
            Add-Log -Log "Invalid path for object '$dn' (skipped)" -Type Warning
            continue
        }
        $description = Get-XmlValue -Xml $obj -XmlPath $Script:DefDescription
        $enabled = Get-XmlValue -Xml $obj -XmlPath $Script:DefEnabled -DefaultValue 'true'
        $samAccountName = Get-XmlValue -Xml $obj -XmlPath $Script:DefSamAccountName
        $perms = Get-XmlValue -Xml $obj -XmlPath $Script:DefPermissions
        if ($perms) {
            if ($Permissions.ContainsKey($dnLc)) {
                $Permissions[$dnLc] += $perms
            } else {
                $Permissions.Add($dnLc, $perms)
            }
        }
        $ps = @{
            'Description' = $description
            'Enabled' = ($enabled -ieq 'true')
            'SamAccountName' = $samAccountName
        }
        $clear = @()
        $parameters = @{}
        foreach ($key in $ps.Keys) {
            if ($ps[$key] -eq $null) {
                continue
            }
            if (($ps[$key] -is [String]) -and ($ps[$key] -eq '')) {
                $clear += $key
            } else {
                $parameters.Add($key, $ps[$key])
            }
        }
        $computer = $null
        try {
            $computer = Get-ADComputer -Identity $dn -Properties 'DistinguishedName' @commonParams
        } catch {
            if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                Add-Log -Log "Error while getting User '$dn' information. $_" -Type Error
                return $false
            }
        }
        $parameters += $commonParams
        try {
            if ($computer) {
                if ($Force) {
                    Add-Log -Log "Updating object '$dn'"
                    if ($clear.Count -gt 0) {
                        $parameters.Add('Clear', $clear)
                    }
                    Set-ADComputer -Identity $dn @parameters
                } else {
                    Add-Log -Log "Object '$dn' already exists"
                }
            } else {
                New-ADComputer -Name $name -Path $path @parameters
                Add-Log -Log "Created object '$dn'"
            }
        } catch {
            Add-Log -Log "Error while restoring object '$dn'. $_" -Type Error
            return $false
        }
    }
    return $true
}

function Restore-Groups {
<#
.SYNOPSIS
    Restores Group objects.
.DESCRIPTION
    Restores Group objects.
.PARAMETER Data
    Data to restore.
.PARAMETER Permissions
    Variable listing permissions (out).
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$false)]
        [System.Xml.XmlElement]$Data = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$Permissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    Add-Log -Log 'Restoring Group Objects' -Type Title2
    if (!$Data) {
        Add-Log -Log 'No Group object to restore'
        return $true
    }
    if (!$Domain) {
        Add-Log -Log 'Active Directory object cannot be null' -Type Error
        return $false
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    if (!$Permissions) {
        $Permissions = @{}
    }
    $groupMembers = @{}
    foreach ($obj in $Data.ChildNodes) {
        $dn = Get-XmlValue -Xml $obj -XmlPath $Script:DefDistinguishedName
        if ([String]::IsNullOrEmpty($dn)) {
            Add-Log -Log 'DistinguishedName cannot be empty (skipped)' -Type Warning
            continue
        }
        $dn += ",$($Domain.DistinguishedName)"
        $dnLc = $dn.ToLower()
        if (![String]::IsNullOrEmpty($Scope) -and !($dnLc.EndsWith($Scope.ToLower()))) {
            Add-Log -Log "Object '$dn' is not in the scope '$Scope' (skipped)" -Type Warning
            continue
        }
        $name = Get-XmlValue -Xml $obj -XmlPath $Script:DefName
        if (($name.Length + 4) -ge $dn.Length) {
            Add-Log -Log "Invalid object's name '$name' (skipped)" -Type Warning
            continue
        }
        $path = Get-PathFromDn -Dn $dn
        if ([String]::IsNullOrEmpty($path)) {
            Add-Log -Log "Invalid path for object '$dn' (skipped)" -Type Warning
            continue
        }
        $description = Get-XmlValue -Xml $obj -XmlPath $Script:DefDescription
        $displayName = Get-XmlValue -Xml $obj -XmlPath $Script:DefDisplayName
        $groupCategory = Get-XmlValue -Xml $obj -XmlPath $Script:DefCategory -DefaultValue 'Security'
        $groupScope = Get-XmlValue -Xml $obj -XmlPath $Script:DefScope -DefaultValue 'Global'
        $samAccountName = Get-XmlValue -Xml $obj -XmlPath $Script:DefSamAccountName
        $members = Get-XmlValue -Xml $obj -XmlPath "$($Script:DefMembers)"
        $perms = Get-XmlValue -Xml $obj -XmlPath $Script:DefPermissions
        if ($perms) {
            if ($Permissions.ContainsKey($dnLc)) {
                $Permissions[$dnLc] += $perms
            } else {
                $Permissions.Add($dnLc, $perms)
            }
        }
        $ps = @{
            'Description' = $description
            'DisplayName' = $displayName
            'GroupCategory' = $groupCategory
            'GroupScope' = $groupScope
            'SamAccountName' = $samAccountName
        }
        $clear = @()
        $parameters = @{}
        foreach ($key in $ps.Keys) {
            if ($ps[$key] -eq $null) {
                continue
            }
            if (($ps[$key] -is [String]) -and ($ps[$key] -eq '')) {
                $clear += $key
            } else {
                $parameters.Add($key, $ps[$key])
            }
        }
        if (!$groupMembers.ContainsKey($dn)) {
            $groupMembers.Add($dn, $members)
        } else {
            Add-Log -Log "Group '$dn' is duplicated (skipped)" -Type Warning
        }
        $group = $null
        try {
            $group = Get-ADGroup -Identity $dn -Properties 'DistinguishedName' @commonParams
        } catch {
            if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                Add-Log -Log "Error while getting object '$dn' information. $_" -Type Error
                return $false
            }
        }
        $parameters += $commonParams
        try {
            if ($group) {
                if ($Force) {
                    Add-Log -Log "Updating object '$dn'"
                    if ($clear.Count -gt 0) {
                        $parameters.Add('Clear', $clear)
                    }
                    Set-ADGroup -Identity $dn @parameters
                } else {
                    Add-Log -Log "Object '$dn' already exists"
                }
            } else {
                New-ADGroup -Name $name -Path $path @parameters
                Add-Log -Log "Created object '$dn'"
            }
        } catch {
            Add-Log -Log "Error while restoring object '$dn'. $_" -Type Error
            return $false
        }
    }
    foreach ($dn in $groupMembers.Keys) {
        if ([String]::IsNullOrEmpty($dn)) {
            continue
        }
        if ($groupMembers[$dn] -eq $null) {
            continue
        }
        if (($groupMembers[$dn] -is [String]) -and ($groupMembers[$dn] -eq '')) {
            try {
                $members = Get-ADGroupMember -Identity $dn @commonParams
                if ($members) {
                    $members = $members.DistinguishedName
                    Remove-ADGroupMember -Identity $dn -Members $members -Confirm:$false @commonParams
                    Add-Log -Log "Removed '$($members -join ''', ''')' from group '$dn'"
                }
            } catch {
                Add-Log -Log "Unable to remove members from group '$dn'. $_" -Type Error
                return $false
            }
            continue
        }
        $members = @()
        foreach ($node in $groupMembers[$dn].ChildNodes) {
            $member = Get-XmlValue -Xml $node -XmlPath '#text'
            $isWellKnown = ((Get-XmlValue -Xml $node -XmlPath $Script:DefIsWellKnown) -ieq 'true')
            if ([String]::IsNullOrEmpty($member)) {
                Add-Log -Log 'Member cannot be empty (skipped)' -Type Warning
                continue
            }
            $f = "samAccountName -eq '$member'"
            if ($isWellKnown) {
                if ($member -imatch '^(\[(Remote|Root)?DomainSID\])-(\d+)$') {
                    if ($Matches[1] -ieq '[DomainSID]') {
                        $member = "$($Domain.DomainSID)-$($Matches[3])"
                    } elseif ($Matches[1] -ieq '[RootDomainSID]') {
                        $member = "$($Domain.RootDomainSID)-$($Matches[3])"
                    } else {
                        Add-Log -Log "Member '$member' is from a remote domain (skipped)" -Type Warning
                        continue
                    }
                }
                $f = "objectSid -eq '$member'"
            }
            try {
                $o = Get-ADObject -Filter $f -Properties 'DistinguishedName' -SearchBase $Scope @commonParams
                $member = $o.DistinguishedName
            } catch {
                Add-Log -Log "Unable to find the group member '$member' under '$Scope'" -Type Error
                return $false
            }
            if (![String]::IsNullOrEmpty($member)) {
                $members += $member
            }
        }
        try {
            Add-ADGroupMember -Identity $dn -Members $members @commonParams
            Add-Log -Log "Added '$($members -join ''', ''')' to group '$dn'"
        } catch {
            Add-Log -Log "Error while adding '$($members -join ''', ''')' to group '$dn'. $_" -Type Error
            return $false
        }
    }
    return $true
}

function Restore-WmiFilters {
<#
.SYNOPSIS
    Restores Wmi Filter objects.
.DESCRIPTION
    Restores Wmi Filter objects.
.PARAMETER Data
    Data to restore.
.PARAMETER Permissions
    Variable listing permissions (out).
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$false)]
        [System.Xml.XmlElement]$Data = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$Permissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    Add-Log -Log 'Restoring WMI Filter Objects' -Type Title2
    if (!$Data) {
        Add-Log -Log 'No WMI Filter object to restore'
        return $true
    }
    if (!$Domain) {
        Add-Log -Log 'Active Directory object cannot be null' -Type Error
        return $false
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    $currentUser = $env:USERNAME
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
        $currentUser = $Credential.UserName
    }
    if ($currentUser.Contains('/')) {
        $currentUser = $currentUser.Split('/')[1]
    }
    if ($currentUser.Contains('@')) {
        $currentUser = $currentUser.Split('@')[0]
    }
    $currentUser += "@$($Domain.DNSRoot)"
    if (!$Permissions) {
        $Permissions = @{}
    }
    $wmiPath = $null
    $wmiObjects = $null
    $currentDate = $null
    try {
        Add-Log -Log 'Collecting existing WMI Filters'
        $wmiPath = "CN=SOM,CN=WMIPolicy,$($Domain.SystemsContainer)"
        $wmiObjects = Get-ADObject -SearchBase $wmiPath `
                                   -Filter { objectClass -eq 'msWMI-Som' } `
                                   -Properties 'distinguishedName', 'msWMI-Name' `
                                   @commonParams
        $currentDate = (Get-Date -ErrorAction Stop).ToUniversalTime().ToString('yyyyMMddhhmmss.ffffff-000')
    } catch {
        Add-Log -Log "Error while listing existing WMI filters. $_" -Type Error
        return $false
    }
    foreach ($obj in $Data.ChildNodes) {
        $name = Get-XmlValue -Xml $obj -XmlPath $Script:DefWmiName
        if ([String]::IsNullOrEmpty($name)) {
            Add-Log -Log 'Name cannot be empty (skipped)' -Type Warning
            continue
        }
        $parm1 = Get-XmlValue -Xml $obj -XmlPath $Script:DefWmiParm1
        $parm2 = Get-XmlValue -Xml $obj -XmlPath $Script:DefWmiParm2
        $author = Get-XmlValue -Xml $obj -XmlPath $Script:DefWmiAuthor -DefaultValue $currentUser
        $showAdvView = Get-XmlValue -Xml $obj -XmlPath $Script:DefShowAdvView -DefaultValue 'true'
        $perms = Get-XmlValue -Xml $obj -XmlPath $Script:DefPermissions
        $filter = $wmiObjects | Where-Object { $_.'msWMI-Name' -ieq $name }
        $parameters = @{
            'msWMI-Name' = $name
            'msWMI-Parm2' = $parm2
            'msWMI-Author' = $author
            'showInAdvancedViewOnly' = $showAdvView.ToUpper()
            'msWMI-ChangeDate' = $currentDate
        }
        if ($parm1 -ne $null) {
            $parameters.Add('msWMI-Parm1', $parm1)
        }
        $dn = $null
        try {
            if ($filter) {
                $dn = $filter.distinguishedName
                if ($Force) {
                    Set-ADObject -Identity $filter.distinguishedName -Replace $parameters @commonParams
                    Add-Log "Updated WMI Filter '$($filter.'msWMI-Name')'"
                } else {
                    Add-Log "WMI Filter '$($filter.'msWMI-Name')' already exists"
                }
                continue
            }
            $guid = "{$([System.Guid]::NewGuid())}"
            $parameters.Add('msWMI-ID', $guid)
            $parameters.Add('msWMI-CreationDate', $currentDate)
            $dn = "CN$($guid),$($wmiPath)"
            New-ADObject -Name $guid `
                         -Type 'msWMI-Som' `
                         -Path $wmiPath `
                         -OtherAttributes $parameters `
                         @commonParams
            Add-Log "Created WMI Filter '$($name)'"
        } catch {
            Add-Log -Log "Error while restoring object '$($name)'. $_" -Type Error
            return $false
        }
        if ($perms) {
            $dnLc = $dn.ToLower()
            if ($Permissions.ContainsKey($dnLc)) {
                $Permissions[$dnLc] += $perms
            } else {
                $Permissions.Add($dnLc, $perms)
            }
        }
    }
    return $true
}

function Restore-MigrationTable {
<#
.SYNOPSIS
    Builds a migration table to be used by restored GPOs.
.DESCRIPTION
    Builds a migration table to be used by restored GPOs.
.PARAMETER Template
    Migration Table template file.
.PARAMETER FilePath
    Save the migration table to this file.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    param (
        [parameter(mandatory=$true)]
        [String]$Template = $null,

        [parameter(mandatory=$true)]
        [String]$FilePath = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    Add-Log -Log 'Creating Migration Table'
    if ([String]::IsNullOrEmpty($Template) -or !(Test-Path -Path $Template)) {
        Add-Log -Log 'Migration table template can not be found' -Type Error
        return $false
    }
    if ([String]::IsNullOrEmpty($FilePath) -or !(Test-Path -Path $FilePath -IsValid)) {
        Add-Log -Log 'Migration table file must be a valid path' -Type Error
        return $false
    }
    if (!(Test-Path -Path $FilePath)) {
        try {
            New-Item -Path $FilePath -ItemType File -Force -ErrorAction Stop | Out-Null
            $FilePath = (Resolve-Path -Path $FilePath -ErrorAction Stop).Path
        } catch {
            Add-Log -Log "Unable to create the file '$FilePath'. $_" -Type Error
            return $false
        }
    }
    try {
        $prop = Get-ItemProperty -Path $Template -ErrorAction Stop
        if ($prop.FullName.StartsWith('\\')) {
            Copy-Item -Path $Template -Destination $Script:TempFolder -ErrorAction Stop
            $Template = Join-Path -Path $Script:TempFolder -ChildPath $prop.Name
        } else {
            $Template = $prop.FullName
        }
        $FilePath = (Resolve-Path -Path $FilePath -ErrorAction Stop).Path
    } catch {
        Add-Log -Log "Unable to resolve path '$FilePath'. $_" -Type Error
        return $false
    }
    $migtable = $null
    $success = $true
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    try {
        $migtable = New-Object -TypeName Microsoft.GroupPolicy.GPMigrationTable($Template)
        foreach ($entry in $migtable.GetEntries()) {
            if (!$entry) {
                continue
            }
            $d = $entry.Destination
            if ([String]::IsNullOrEmpty($d)) {
                $arr = $entry.Source.Split('@')
                if ((@($arr).Count -gt 1) -and ($arr[0] -ine $Domain.DNSRoot)) {
                    $d = $arr[0] + "@$($Domain.DNSRoot)"
                } else {
                    #$d = $arr[0] + "@$($Domain.DNSRoot)"
                }
            } elseif ($d -imatch '^(\[(Remote|Root)?DomainSID\])-(\d+)$') {
                $parameters = @{
                    'Properties' = @('objectSid', 'SamAccountName')
                    'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
                    'Server' = $Server
                }
                if ([String]::IsNullOrEmpty($Server)) {
                    $parameters['Server'] = $Domain.PDCEmulator
                }
                if ($Credential) {
                    $parameters.Add('Credential', $Credential)
                }
                $sid = [String]::Empty
                $suffix = $null
                if ($Matches[1] -ieq '[RemoteDomainSID]') {
                    Add-Log -Log "Item '$($entry.Source)' is from a remote domain that can't be translated" `
                            -Type Warning
                } elseif ($Matches[1] -ieq '[RootDomainSID]') {
                    $sid = "$($Domain.RootDomainSID)-$($Matches[3])"
                    $parameters['Server'] = $Domain.RootPDCEmulator
                    $suffix = $Domain.Forest
                } else {
                    $sid = "$($Domain.DomainSID)-$($Matches[3])"
                    $suffix = $Domain.DNSRoot
                }
                if (![String]::IsNullOrEmpty($sid)) {
                    $d = Convert-SidToString -Sid $sid -Server $Server -Credential $Credential -Domain $Domain
                    if ([String]::IsNullOrEmpty($d)) {
                        $d = [String]::Empty
                    } else {
                        $d = "$($d)@$($suffix)"
                    }
                }
            } else {
                $sid = Convert-StringToSid -String $d -Server $Server -Credential $Credential -Domain $Domain
                if ($sid) {
                    $d = Convert-SidToString -Sid $sid -Server $Server -Credential $Credential -Domain $Domain
                    if ([String]::IsNullOrEmpty($d)) {
                        try {
                            $d = $sid.Translate([System.Security.Principal.NTAccount])
                            if ($d -and ![String]::IsNullOrEmpty($d.Value) -and $d.Value.Contains('\')) {
                                $d = $d.Value.Split('\')[1]
                            }
                        } catch {
                            if ("$sid" -ine 'S-1-5-32-547') { # Power Users
                                Add-Log -Log "Unable to translate '$($entry.Source)'. $_" -Type Warning
                            }
                            $d = [String]::Empty
                        }
                    }
                }
            }
            if ([String]::IsNullOrEmpty($d)) {
                $migtable.UpdateDestination($entry.Source) | Out-Null
            } else {
                #if ($entry.Source -ine $d) {
                    Add-Log -Log "Mapped '$($entry.Source)' to '$($d)'"
                #}
                $migtable.UpdateDestination($entry.Source, $d) | Out-Null
            }
        }
        if ($success) {
            $migtable.Save($FilePath)
            Add-Log -Log "Migration table saved to '$FilePath'"
        }
    } catch {
        Add-Log -Log "Unable to build the migration table. $_" -Type Error
        $success = $false
    }
    $migtable = $null
    return $success
}

function Restore-Admx {
<#
.SYNOPSIS
    Restores ADMX and ADML files to the Central Store.
.DESCRIPTION
    Restores ADMX and ADML files to the Central Store.
.PARAMETER BackupFolder
    Path to backup folder.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$false)]
        [String]$BackupFolder = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    Add-Log -Log 'Restoring ADMX and ADML files to the Central Store' -Type Title2
    if (!$Domain) {
        Add-Log -Log 'Active Directory object cannot be null' -Type Error
        return $false
    }
    $BackupFolder = Join-Path -Path $BackupFolder -ChildPath $Script:AdmxFolderName
    if ([String]::IsNullOrEmpty($BackupFolder) -or !(Test-Path -Path $BackupFolder)) {
        Add-Log -Log "No file to restore"
        return $true
    }
    Add-Log -Log "Backup folder: '$BackupFolder'"
    $items = $null
    try {
        $items = Get-ChildItem -Path $BackupFolder -Force -ErrorAction Stop
        if (!$items) {
            Add-Log -Log 'No file to restore'
            return $true
        }
        $items = $items | Sort-Object -Property FullName
    } catch {
        Add-Log -Log "Error while listing files in folder '$BackupFolder'. $_" -Type Error
        return $false
    }
    if ([String]::IsNullOrEmpty($Server)) {
        $Server = $Domain.PDCEmulator
    }
    $parameters = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($Credential) {
        $parameters.Add('Credential', $Credential)
    }
    $drive = $null
    $success = $true
    try {
        $dest = "\\$($Server)\SYSVOL\$($Domain.DNSRoot)\Policies"
        $drive = New-PSDrive -Name 'admx' -PSProvider FileSystem -Root $dest @parameters
        Add-Log -Log "Mapped drive 'admx:\' to '$dest'"
        $dest = 'admx:\PolicyDefinitions'
        if (!(Test-Path -Path $dest)) {
            $folder = New-Item -Path $dest -ItemType Directory -ErrorAction Stop
            Add-Log -Log "Created folder '$($folder.FullName)'"
        }
        foreach ($item in $items) {
            if (!$item) {
                continue
            }
            try {
                Add-Log -Log "Copying item '$($item.FullName)' to '$dest'"
                Copy-Item -Path $item.FullName `
                          -Destination $dest `
                          -Force:$true `
                          -Recurse:$item.PSIsContainer `
                          -ErrorAction Stop
            } catch {
                if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ResourceExists) {
                    Add-Log -Log "Unable to copy item '$($item.FullName)' to '$dest'. $_" -Type Error
                    $success = $false
                    break
                }
                Add-Log -Log "Item '$($item.Name)' already exists in '$dest' (skipped)"
            }
        }
    } catch {
        Add-Log -Log "Error while updating Central Store. $_" -Type Error
        $success = $false
    } finally {
        try {
            $drive | Remove-PSDrive -Force -ErrorAction Stop
        } catch {
            Add-Log -Log "Unable to remove the PSDrive. $_" -Type Error
            $success = $false
        }
    }
    return $success
}

function Restore-GroupPolicies {
<#
.SYNOPSIS
    Restores Group Policy objects and WMI Filters.
.DESCRIPTION
    Restores Group Policy objects and WMI Filters.
.PARAMETER Data
    Data to restore.
.PARAMETER BackupFolder
    Path to GPOs backup folder.
.PARAMETER All
    If true, restores all GPOs, not only GPOs with a link.
.PARAMETER WmiFilter
    If true, WMI Filter are applied to GPOs.
.PARAMETER Permissions
    Variable listing permissions (out).
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    Hashtable with all GPOs to link or true if no GPO to import.
    False if an error occurs or null if no GPO to link.
#>
    param(
        [Parameter(Mandatory=$false)]
        [System.Xml.XmlElement]$Data = $null,

        [Parameter(Mandatory=$false)]
        [String]$BackupFolder = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$All = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$WmiFilter = $false,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$Permissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    Add-Log -Log 'Restoring Group Policy Objects' -Type Title2

    # Checks requirements.
    if (!$Data) {
        Add-Log -Log 'No Group Policy object to restore'
        return $true
    }
    if (!$Domain) {
        Add-Log -Log 'Active Directory object cannot be null' -Type Error
        return $false
    }
    if ([String]::IsNullOrEmpty($BackupFolder) -or !(Test-Path -Path $BackupFolder)) {
        Add-Log -Log "Backup folder '$BackupFolder' not found" -Type Error
        return $false
    }
    $tmpFolder = $BackupFolder
    try {
        $files = Get-ChildItem -Path $tmpFolder -File -Recurse -ErrorAction Stop |
                 Where-Object { $_.Attributes -band [System.IO.FileAttributes]::ReadOnly }
        if ($files) {
            $tmpFolder = $Script:TempFolder
            Add-Log -Log "Copying folder '$BackupFolder' to '$tmpFolder' and removing ReadOnly Attribute from files"
            Copy-Item -Path "$BackupFolder\*" -Destination $tmpFolder -Recurse -Force -ErrorAction Stop
            Get-ChildItem -Path $tmpFolder -File -Recurse -Force -ErrorAction Stop |
            Where-Object { $_.Attributes -band [System.IO.FileAttributes]::ReadOnly } |
            ForEach-Object {
                $_.Attributes = ($_.Attributes -bxor [System.IO.FileAttributes]::ReadOnly)
            }
        }
    } catch {
        Add-Log -Log "Unable to prepare GPO backup files. $_" -Type Error
        return $false
    }
    Add-Log -Log "Backup folder: '$tmpFolder'"
    $migtable = $null
    try {
        $files = Get-ChildItem -Path $tmpFolder -File -Filter $Script:MigrationTableFile -ErrorAction Stop
        if ($files) {
            if ($files -isnot [System.IO.FileInfo]) {
                Add-Log -Log "Multiple Migration Tables found in folder '$tmpFolder'" -Type Error
                return $false
            }
            $migtable = Join-Path -Path $Script:TempFolder -ChildPath "$($Domain.DNSRoot).migtable" -ErrorAction Stop
            if (!(Restore-MigrationTable -Template $files.FullName -FilePath $migtable -Domain $Domain)) {
                return $false
            }
            Add-Log -Log "Migration Table: '$migtable'"
        } else {
            Add-Log -Log 'No Migration Table found'
        }
    } catch {
        Add-Log -Log "Error while listing files from folder '$tmpFolder'. $_" -Type Error
        return $false
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    if (!$Permissions) {
        $Permissions = @{}
    }

    # Imports GPOs and collects GPO Links information.
    $hOuLinks = @{}
    $importedGpoNames = @{}
    foreach ($obj in ($Data.ChildNodes | Sort-Object -Property $Script:DefName)) {
        $gpoStatus = Get-XmlValue -Xml $obj -XmlPath $Script:DefGpoStatus -DefaultValue 'AllSettingsEnabled'
        $links = Get-XmlValue -Xml $obj -XmlPath "$($Script:DefLinks)"
        $name = Get-XmlValue -Xml $obj -XmlPath $Script:DefName
        $targetName = Get-XmlValue -Xml $obj -XmlPath $Script:DefTargetName
        if ([String]::IsNullOrEmpty($name)) {
            Add-Log -Log 'Name of a GPO cannot be empty' -Type Error
            return $false
        }
        if ([String]::IsNullOrEmpty($targetName)) {
            $targetName = $name
        }
        $perms = Get-XmlValue -Xml $obj -XmlPath $Script:DefPermissions
        if (($links -eq $null) -and !$All) {
            continue
        }

        # Imports GPO.
        $gpo = $null
        try {
            $gpo = Get-XGPO -Name $targetName -Domain $Domain.DNSRoot @commonParams
        } catch {
            if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                Add-Log -Log "Error while getting GPO '$targetName'. $_" -Type Error
                return $false
            }
        }
        if (!$gpo -or $Force) {
            $gpoParams = $commonParams + @{
                'BackupGpoName' = $name
                'Domain' = $Domain.DNSRoot
                'Path' = $tmpFolder
                'TargetName' = $targetName
                'CreateIfNeeded' = $true
            }
            if (![String]::IsNullOrEmpty($migtable)) {
                $gpoParams.Add('MigrationTable', $migtable)
            }
            $gpo = $null
            try {
                $gpo = Import-XGPO @gpoParams
                Add-Log -Log "Imported GPO '$($gpo.DisplayName)'"
            } catch {
                Add-Log -Log "Error while importing GPO '$name'. $_" -Type Error
                return $false
            }
        } else {
            Add-Log -Log "GPO '$($gpo.DisplayName)' already exists"
        }
        if (!$importedGpoNames.ContainsKey("$($gpo.DisplayName)".ToLower())) {
            $importedGpoNames.Add("$($gpo.DisplayName)".ToLower(), $null)
        }
        $dn = $gpo.Path
        $dnLc = $dn.ToLower()
        if ($perms) {
            if ($Permissions.ContainsKey($dnLc)) {
                $Permissions[$dnLc] += $perms
            } else {
                $Permissions.Add($dnLc, $perms)
            }
        }

        # Collects GPO Links information.
        if ($links -eq $null) {
            continue
        }
        if ($links -eq '') {
            #Add-Log -Log 'Removing all GPO Links' -Type Warning
            #TODO: Remove all links?
            continue
        }
        foreach ($link in $links.ChildNodes) {
            $enabled = Get-XmlValue -Xml $link -XmlPath "$($Script:DefEnabled)" -DefaultValue 'true'
            $enforced = Get-XmlValue -Xml $link -XmlPath "$($Script:DefEnforced)" -DefaultValue 'false'
            $order = Get-XmlValue -Xml $link -XmlPath "$($Script:DefOrder)"
            $container = Get-XmlValue -Xml $link -XmlPath "$($Script:DefContainer)"
            if ([String]::IsNullOrEmpty($container)) {
                $container = "$($Domain.DistinguishedName)"
            } else {
                $container += ",$($Domain.DistinguishedName)"
            }
            $containerLc = $container.ToLower()
            if ($enabled -ieq 'true') {
                $enabled = [Microsoft.GroupPolicy.EnableLink]::Yes
            } else {
                $enabled = [Microsoft.GroupPolicy.EnableLink]::No
            }
            if ($enforced -ieq 'true') {
                $enforced = [Microsoft.GroupPolicy.EnforceLink]::Yes
            } else {
                $enforced = [Microsoft.GroupPolicy.EnforceLink]::No
            }
            try {
                $order = [UInt32]$order
            } catch {
                $s = "Link order (container '$container') must be a positive integer and has been set to " +
                     "'1' instead of '$order'"
                $order = 1
                Add-Log -Log $s -Type Warning
            }
            if (!$hOuLinks.ContainsKey($containerLc)) {
                $hOuLinks.Add($containerLc, @{})
            }
            if ($hOuLinks[$containerLc].ContainsKey($order)) {
                $order = [UInt32](($hOuLinks[$containerLc].Keys | Sort-Object -Descending)[0]) + 1
                $s = "Link order (container '$container') has been changed to '$order' " +
                     'because another GPO has the same link order'
                Add-Log -Log $s -Type Warning
            }
            $hOuLinks[$containerLc][$order] = @($gpo, $enabled, $enforced)
        }
    }
    if (!$WmiFilter) {
        return $hOuLinks
    }

    # Applying WMI Filters to GPOs.
    Add-Log -Log 'Applying WMI Filters to GPOs' -Type Title2
    $wmiObjects = $null
    $wmiPath = "CN=SOM,CN=WMIPolicy,$($Domain.SystemsContainer)"
    try {
        $wmiObjects = Get-ADObject -SearchBase $wmiPath `
                                   -Filter { objectClass -eq 'msWMI-Som' } `
                                   -Properties 'distinguishedName', 'msWMI-Name' `
                                   @commonParams
    } catch {
        Add-Log -Log "Error while listing WMI Filters in $($wmiPath). $_" -Type Error
        return $false
    }
    if (!$wmiObjects) {
        Add-Log -Log 'No WMI Filter found in Active Directory'
        return $hOuLinks
    }
    $backups = Get-ChildItem -Path $tmpFolder -File -Filter 'Backup.xml' -Recurse -ErrorAction SilentlyContinue
    foreach ($backup in $backups) {
        if (!$backup) {
            continue
        }
        $gpoName = $null
        $filterName = $null
        try {
            [Xml]$x = Get-Content -Path $backup.FullName -ErrorAction Stop
            $node = $x.GroupPolicyBackupScheme.GroupPolicyObject.GroupPolicyCoreSettings
            $gpoName = $node.DisplayName.'#cdata-section'
            $filterName = $node.WMIFilterName.'#cdata-section'
        } catch {
        }
        if ([String]::IsNullOrEmpty($gpoName)) {
            Add-Log -Log "Unable to find the GPO's name from file '$($backup.FullName)'" -Type Error
            return $false
        }
        if ([String]::IsNullOrEmpty($filterName)) {
            continue
        }
        if (!$importedGpoNames.ContainsKey($gpoName.ToLower())) {
            continue
        }
        $filter = $wmiObjects | Where-Object { $_.'msWMI-Name' -ieq $filterName }
        if (!$filter) {
            Add-Log -Log "Cannot apply WMI Filter '$($filterName)' to GPO '$($gpoName)' (filter not found)" `
                    -Type Warning
            continue
        }
        $gpo = $null
        try {
            $gpo = Get-XGPO -Name $gpoName -Domain $Domain.DNSRoot @commonParams
        } catch {
            if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                Add-Log -Log "Error while getting GPO '$targetName'. $_" -Type Error
                return $false
            }
            continue
        }
        try {
            if ($gpo.WmiFilter -and ($gpo.WmiFilter.Name -ieq $filterName) -and !$Force) {
                Add-Log -Log "WMI Filter '$($filterName)' already applied to GPO '$($gpoName)'"
            } else {
                Set-ADObject -Identity $gpo.Path `
                             -Replace @{ gPCWQLFilter = "[$($Domain.DNSRoot);$($filter.Name);0]" } `
                             @commonParams
                Add-Log -Log "WMI Filter '$($filterName)' applied to GPO '$($gpoName)'"
            }
        } catch {
            Add-Log -Log "Error while applying WMI Filter '$($filterName)' to GPO '$($gpoName)'. $_" -Type Error
            return $false
        }
    }
    return $hOuLinks
}

function Restore-GroupPolicyLinks {
<#
.SYNOPSIS
    Links Group Policy objects to containers.
.DESCRIPTION
    Links Group Policy objects to containers.
.PARAMETER Links
    GPO Links as Hashtable:
    [$containerLc][$order] = @($gpo, $enabled, $enforced).
.PARAMETER GpoLinks
    Defines how GPOs are linked.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Collections.Hashtable]$Links = $null,

        [Parameter(Mandatory=$false)]
        [ValidateSet('DontLink', 'LinkDisabled', 'LinkEnabled')]
        [String]$GpoLinks = 'DontLink',

        [Parameter(Mandatory=$false)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    if (($GpoLinks -ieq 'DontLink') -and !$Force) {
        return $true
    }
    Add-Log -Log 'Configuring GPOs'' Links' -Type Title2
    if (!$Links -or ($Links.Count -le 0)) {
        Add-Log 'No GPO Link found'
        return $true
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    foreach ($container in $Links.Keys) {
        $orders = @() + ([UInt32[]]@($Links[$container].Keys) | Sort-Object)
        for ($i = 0; $i -lt $orders.Count; $i++) {
            $order = $orders[$i]
            $gpo = $Links[$container][$order][0]
            if (!$gpo) {
                Add-Log 'GPO cannot be null (skipped)' -TYpe Warning
                continue
            }
            $name = $gpo.DisplayName
            $guid = $gpo.Id
            $enabled = $Links[$container][$order][1]
            $enforced = $Links[$container][$order][2]
            $linked = $false
            try {
                $linked = Get-XGPInheritance -Target $container -Domain $Domain.DNSRoot @commonParams
                $linked = ($linked -and $linked.GpoLinks -and ($guid -iin @($linked.GpoLinks.GpoId)))
            } catch {
                if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                    Add-Log -Log "Error while getting GPO link information for '$container'. $_" -Type Error
                    return $false
                }
                $linked = $false
            }
            if (!$linked -and ($GpoLinks -ieq 'DontLink')) {
                Add-Log -Log "GPO not linked: '$($name)'"
                continue
            }
            if ($linked -and !$Force) {
                $s = "GPO '$($name)' already linked to '$($container)'"
                if ($GpoLinks -ieq 'DontLink') {
                    $s += ': use -Force parameter to remove the link'
                }
                Add-Log -Log $s
                continue
            }
            if ($linked -and ($GpoLinks -ieq 'DontLink')) {
                try {
                    Remove-XGPLink -Guid $guid -Target $container -Domain $Domain.DNSRoot @commonParams | Out-Null
                    Add-Log "Removed GPO '$name' link from '$container'"
                    continue
                } catch {
                    Add-Log "Error while removing GPO '$name' link from '$container'. $_" -Type Error
                    return $false
                }
            }
            if (($enabled -ieq [Microsoft.GroupPolicy.EnableLink]::Yes) -and ($GpoLinks -ieq 'LinkDisabled')) {
                $enabled = [Microsoft.GroupPolicy.EnableLink]::No
            }
            $parameters = $commonParams + @{
                'Guid' = $guid
                'Target' = $container.Replace('/', '\/')
                'LinkEnabled' = $enabled
                'Order' = $order
                'Enforced' = $enforced
                'Domain' = $Domain.DNSRoot
            }
            try {
                if (!$linked) {
                    New-XGPLink @parameters | Out-Null
                } else {
                    Set-XGPLink @parameters | Out-Null
                }
                $s = "GPO '$($name)' linked to '$($container)' (priority $order, link "
                if ($parameters['LinkEnabled'] -ieq [Microsoft.GroupPolicy.EnableLink]::Yes) {
                    $s += 'enabled'
                } else {
                    $s += 'disabled'
                }
                if ($parameters['Enforced'] -ieq [Microsoft.GroupPolicy.EnforceLink]::Yes) {
                    $s += ' and enforced'
                }
                $s += ')'
                Add-Log -Log $s
            } catch {
                Add-Log -Log "Error while linking GPO '$name' to '$container'. $_" -Type Error
                return $false
            }
        }
    }
    return $true
}

function Restore-Permissions {
<#
.SYNOPSIS
    Restores permissions.
.DESCRIPTION
    Restores permissions.
.PARAMETER Data
    Data to restore.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$Data = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    Add-Log -Log 'Restoring Permissions' -Type Title2
    if (!$Data) {
        Add-Log -Log 'No permissions to restore'
        return $true
    }
    if (!$Domain) {
        Add-Log -Log 'Active Directory object cannot be null' -Type Error
        return $false
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    $success = $true
    foreach ($dnLc in ($Data.Keys | Sort-Object)) {
        if ([String]::IsNullOrEmpty($dnLc)) {
            continue
        }
        foreach ($xml in $Data[$dnLc]) {
            $perms = Get-AclFromXml -Path $dnLc -XmlNode $xml -Domain $Domain @commonParams
            if (!$perms) {
                continue
            }
            $dn = $perms.Path
            $o = $null
            try {
                $o = Get-ADObject -Identity $dn -Properties 'distinguishedName', 'ntSecurityDescriptor' @commonParams
            } catch {
                if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                    Add-Log -Log "Error while getting object '$dn' information. $_" -Type Error
                    return $false
                }
            }
            if (!$o) {
                Add-Log -Log "Object '$dn' not found (skipped)" -Type Warning
                continue
            }
            try {
                $acl = $o.ntSecurityDescriptor
                $s = "Adding permissions to object '$dn':"
                foreach ($ace in @($perms.AceList | Sort-Object -Property 'IdentityReference')) {
                    $acl.AddAccessRule($ace)
                    $id = "$($ace.IdentityReference)"
                    foreach ($name in $Script:NameToSidCache.Keys) {
                        if ($id -ieq $Script:NameToSidCache[$name]) {
                            $id = $name
                            break
                        }
                    }
                    $s += [Environment]::NewLine + "Identity '$($id)': "
                    $s += "$($ace.AccessControlType) $($ace.ActiveDirectoryRights)"
                    $guid = $null
                    try {
                        $guid = [System.Guid]$ace.ObjectType
                    } catch {
                        $guid = $null
                    }
                    if (($guid -is [System.Guid]) -and ($guid -ne [System.Guid]::Empty)) {
                        if ($Script:SchemaAtributes.ContainsKey("$guid".ToLower())) {
                            $s += " ($($Script:SchemaAtributes[$guid.ToString().ToLower()]))"
                        } else {
                            $s += " ($guid)"
                        }
                    }
                    $guid = $null
                    try {
                        $guid = [System.Guid]$ace.InheritedObjectType
                    } catch {
                        $guid = $null
                    }
                    if (($guid -is [System.Guid]) -and ($guid -ne [System.Guid]::Empty)) {
                        if ($Script:SchemaAtributes.ContainsKey("$guid".ToLower())) {
                            $s += " on object $($Script:SchemaAtributes[$guid.ToString().ToLower()])"
                        } else {
                            $s += " on object $guid"
                        }
                    }
                }
                Add-Log -Log $s
                Set-ADObject -Identity $dn -Replace @{ ntSecurityDescriptor = $acl } @commonParams
            } catch {
                Add-Log -Log "Error while updating permissions on '$dn'. $_" -Type Error
                $success = $false
            }
        }
    }
    return $success
}

function Restore-Objects {
<#
.SYNOPSIS
    Restores objects.
.DESCRIPTION
    Restores objects.
.PARAMETER Xml
    Xml document built from Xml settings file.
.PARAMETER OU
    If true, restore Organizational Unit objects.
.PARAMETER User
    If true, restore User objects.
.PARAMETER Computer
    If true, restore Computer objects.
.PARAMETER Group
    If true, restore Group objects.
.PARAMETER GPO
   Restores Group Policy objects.
.PARAMETER GpoLinks
    Defines how GPOs are linked.
.PARAMETER Permissions
    If true, permissions are restored.
.PARAMETER RedirectContainers
    If true, backs up redirection of containers.
.PARAMETER WmiFilters
    If true, Wmi Filters are restored.
.PARAMETER Admx
    If true, ADMX and ADML files are restored to the Central Store.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$Xml = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$OU = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$User = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$Computer = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$Group = $false,

        [Parameter(Mandatory=$false)]
        [ValidateSet('All', 'LinkedOnly', 'None')]
        [String]$GPO = 'None',

        [Parameter(Mandatory=$false)]
        [ValidateSet('DontLink', 'LinkDisabled', 'LinkEnabled')]
        [String]$GpoLinks = 'DontLink',

        [Parameter(Mandatory=$false)]
        [Switch]$Permissions = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$RedirectContainers = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$WmiFilter = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$Admx = $false,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )
    if ($Xml -isnot [System.Xml.XmlDocument]) {
        Add-Log -Log 'Xml cannot be null' -Type Error
        return $false
    }
    $xmlParams = $null
    $xmlData = $null
    try {
        $xmlData = $Xml.$($Script:DefConfiguration)
        $xmlParams = $xmlData.$($Script:DefParameters)
    } catch {
        $xmlParams = $null
        $xmlData = $null
    }
    if ($xmlData -isnot [System.Xml.XmlElement]) {
        Add-Log -Log "Xml node '$($Script:DefConfiguration)' not found" -Type Error
        return $false
    }
    if ($xmlParams -isnot [System.Xml.XmlElement]) {
        Add-Log -Log "Xml node '$($Script:DefConfiguration)/$($Script:DefParameters)' not found" -Type Error
        return $false
    }
    $gpoBackupFolder = Get-XmlValue -Xml $xmlParams -XmlPath "$($Script:DefGpoBackupFolder)"
    if (![String]::IsNullOrEmpty($gpoBackupFolder)) {
        try {
            if (![System.IO.Path]::IsPathRooted($gpoBackupFolder)) {
                $gpoBackupFolder = Join-Path -Path $Script:BaseFolder -ChildPath $gpoBackupFolder -ErrorAction Stop
            }
            $prop = Get-ItemProperty -Path $gpoBackupFolder -ErrorAction Stop
            $gpoBackupFolder = $prop.FullName
        } catch {
            $gpoBackupFolder = $null
        }
    }
    if (!$Domain) {
        Add-Log -Log 'Domain can not be null' -Type Error
        return $false
    }
    $success = $true
    $parameters = @{
        'Domain' = $Domain
        'Server' = $Server
        'Scope' = $Scope
        'Credential' = $Credential
        'Force' = $Force
    }
    try {
        $xmlData.ChildNodes.Name | Out-Null
    } catch {
        Add-Log -Log 'No data to restore' -Type Error
        return $false
    }
    $perms = @{}
    if ($OU) {
        $d = Get-XmlValue -Xml $xmlData -XmlPath "$($Script:DefOus)"
        if ($d) {
            $success = Restore-OrganizationalUnits -Data $d `
                                                   -Permissions $perms `
                                                   -RedirectContainer:$RedirectContainers `
                                                   @parameters
            if (!$success) {
                return $false
            }
        }
    }
    if ($User) {
        $d = Get-XmlValue -Xml $xmlData -XmlPath "$($Script:DefUsers)"
        if ($d) {
            $success = Restore-Users -Data $d -Permissions $perms @parameters
            if (!$success) {
                return $false
            }
        }
    }
    if ($Computer) {
        $d = Get-XmlValue -Xml $xmlData -XmlPath "$($Script:DefComputers)"
        if ($d) {
            $success = Restore-Computers -Data $d -Permissions $perms @parameters
            if (!$success) {
                return $false
            }
        }
    }
    if ($Group) {
        $d = Get-XmlValue -Xml $xmlData -XmlPath "$($Script:DefGroups)"
        if ($d) {
            $success = Restore-Groups -Data $d -Permissions $perms @parameters
            if (!$success) {
                return $false
            }
        }
    }
    if ($WmiFilter) {
        $d = Get-XmlValue -Xml $xmlData -XmlPath "$($Script:DefWmiFilters)"
        if ($d) {
            $success = Restore-WmiFilters -Data $d -Permissions $perms @parameters
            if (!$success) {
                return $false
            }
        }
    }
    if ($Admx) {
        $success = Restore-Admx -BackupFolder $gpoBackupFolder @parameters
        if (!$success) {
            return $false
        }
    }
    if ($GPO -ine 'None') {
        $d = Get-XmlValue -Xml $xmlData -XmlPath "$($Script:DefGroupPolicies)"
        if ($d) {
            $all = ($GPO -ieq 'All')
            $ret = Restore-GroupPolicies -Data $d `
                                         -BackupFolder $gpoBackupFolder `
                                         -All:$all `
                                         -WmiFilter:$WmiFilter `
                                         -Permissions $perms `
                                         @parameters
            if (($ret -eq $false) -or ($ret -eq $null)) {
                return $false
            }
            $success = Restore-GroupPolicyLinks -Links $ret -GpoLinks $GpoLinks @parameters
            if (!$success) {
                return $false
            }
        }
    }
    if ($Permissions) {
        $success = Restore-Permissions -Data $perms @parameters
        if (!$success) {
            return $false
        }
    }
    return $success
}

#endregion

#region Backup

function Backup-OrganizationalUnits {
<#
.SYNOPSIS
    Backs up Organizational Unit objects.
.DESCRIPTION
    Backs up Organizational Unit objects.
.PARAMETER XmlDoc
    Xml Document.
.PARAMETER XmlNode
    Xml Node.
.PARAMETER DefaultPermissions
    Default permissions on AD objects.
.PARAMETER RedirectContainers
    Backs up redirection of containers.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$XmlDoc = $null,

        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$XmlNode = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$DefaultPermissions = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$RedirectContainers = $false,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    Add-Log -Log 'Backing up Organizational Unit objects' -Type Title2
    $ps = @{
        'Description' = $Script:DefDescription
        'DistinguishedName' = $Script:DefDistinguishedName
        'gPOptions' = $Script:DefGpoInheritanceBlocked
        'Name' = $Script:DefName
        'ProtectedFromAccidentalDeletion' = $Script:DefProtectedFromDel
    }
    $permissions = $null
    if ($DefaultPermissions) {
        $ps.Add('ntSecurityDescriptor', $Script:DefPermissions)
        $permissions = $DefaultPermissions['OrganizationalUnit']
    }
    $parameters = @{
        'Properties' = @($ps.Keys)
        'ResultSetSize' = $null
        'SearchScope' = [Microsoft.ActiveDirectory.Management.ADSearchScope]::Subtree
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Scope)) {
        $parameters.Add('SearchBase', $Scope)
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if ($Credential) {
        $parameters.Add('Credential', $Credential)
    }
    $objs = $null
    try {
        $objs = Get-ADOrganizationalUnit -Filter '*' @parameters
        Add-Log -Log "Number of objects found: $(@($objs).Count)"
    } catch {
        Add-Log -Log "Error while listing Active Directory objects. $_" -Type Error
        return $false
    }
    $hContainers = Get-SystemContainer -Domain $Domain -Server $Server -Credential $Credential
    $xs = $xmlDoc.CreateNode('element', $Script:DefOus, $null)
    $xmlNode.AppendChild($xs) | Out-Null
    if ($RedirectContainers) {
        $ps.Add('SystemContainer', $Script:DefSystemContainer)
    }
    foreach ($obj in $objs) {
        if (!$obj) {
            continue
        }
        Add-Log -Log "Backing up object '$($obj.DistinguishedName)'"
        $x = $xmlDoc.CreateNode('element', $Script:DefOu, $null)
        $xs.AppendChild($x) | Out-Null
        $systemContainer = @()
        if ($RedirectContainers) {
            foreach ($containerName in $hContainers.Keys) {
                $dn = $hContainers[$containerName]
                if ($obj.DistinguishedName -ieq $dn) {
                    $systemContainer += $containerName
                }
            }
            if (@($systemContainer).Count -gt 0) {
                $obj | Add-Member -MemberType NoteProperty -Name 'SystemContainer' -Value $systemContainer -Force
            }
        }
        foreach ($p in $ps.keys) {
            if ([String]::IsNullOrEmpty($ps[$p])) {
                continue
            }
            $parameters = @{
                'XmlDoc' = $XmlDoc
                'NodeName' = $ps[$p]
                'Data' = $obj.$p
                'DefaultPermissions' = $permissions
                'Domain' = $Domain
            }
            if (![String]::IsNullOrEmpty($Server)) {
                $parameters.Add('Server', $Server)
            }
            if ($Credential) {
                $parameters.Add('Credential', $Credential)
            }
            $node = Create-SettingsXmlNode @parameters
            if ($node) {
                $x.AppendChild($node) | Out-Null
            }
            if ($node -eq $false) {
                return $false
            }
        }
    }
    return $true
}

function Backup-Users {
<#
.SYNOPSIS
    Backs up User objects.
.DESCRIPTION
    Backs up User objects.
.PARAMETER XmlDoc
    Xml Document.
.PARAMETER XmlNode
    Xml Node.
.PARAMETER DefaultPermissions
    Default permissions on AD objects.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$XmlDoc = $null,

        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$XmlNode = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$DefaultPermissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    Add-Log -Log 'Backing up User objects' -Type Title2
    $ps = @{
        'AccountExpirationDate' = $Script:DefAccountExpirationDate
        'AccountNotDelegated' = $Script:DefAccountNotDelegated
        'AdminCount' = $null
        'AllowReversiblePasswordEncryption' = $Script:DefAllowRevPwdEncryption
        'CannotChangePassword' = $Script:DefCannotChangePassword
        'Description' = $Script:DefDescription
        'DisplayName' = $Script:DefDisplayName
        'DistinguishedName' = $Script:DefDistinguishedName
        'Enabled' = $Script:DefEnabled
        'GivenName' = $Script:DefGivenName
        'Initials' = $Script:DefInitials
        #'KerberosEncryptionType' = $Script:DefKerberosEncryptionType
        'Name' = $Script:DefName
        'PasswordNeverExpires' = $Script:DefPasswordNeverExpires
        'PasswordNotRequired' = $Script:DefPasswordNotRequired
        'ProtectedFromAccidentalDeletion' = $Script:DefProtectedFromDel
        'SamAccountName' = $Script:DefSamAccountName
        'SID' = $null
        'SmartcardLogonRequired' = $Script:DefSmartcardLogonRequired
        'TrustedForDelegation' = $Script:DefTrustedForDelegation
        'userAccountControl' = $null
        'UserPrincipalName' = $Script:DefUserPrincipalName
    }
    if ($DefaultPermissions) {
        $ps.Add('ntSecurityDescriptor', $Script:DefPermissions)
    }
    $defParams = @{
        'ResultSetSize' = $null
        'SearchScope' = [Microsoft.ActiveDirectory.Management.ADSearchScope]::Subtree
    }
    if (![String]::IsNullOrEmpty($Scope)) {
        $defParams.Add('SearchBase', $Scope)
    }
    $parameters = $defParams + @{
        'Properties' = @($ps.Keys)
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    $objs = $null
    $ousDN = $null
    try {
        $ousDN = Get-ADOrganizationalUnit -Filter '*' -Properties 'DistinguishedName' @defParams @commonParams
        $ousDN = @($ousDN.DistinguishedName)
        $objs = Get-ADUser -Filter '*' @parameters @commonParams
        Add-Log -Log "Number of objects found: $(@($objs).Count)"
    } catch {
        Add-Log -Log "Error while listing Active Directory objects. $_" -Type Error
        return $false
    }
    $xs = $xmlDoc.CreateNode('element', $Script:DefUsers, $null)
    $xmlNode.AppendChild($xs) | Out-Null
    foreach ($obj in $objs) {
        if (!$obj) {
            continue
        }
        $path = $obj.DistinguishedName -ireplace '^.+?,(CN|OU.+)', '$1'
        if ($path -inotin $ousDN) {
            Add-Log -Log "Object '$($obj.DistinguishedName)' not in the scope (skipped)"
            continue
        }
        Add-Log -Log "Backing up object '$($obj.DistinguishedName)'"
        $x = $xmlDoc.CreateNode('element', $Script:DefUser, $null)
        $xs.AppendChild($x) | Out-Null
        $permissions = $null
        if ($DefaultPermissions) {
            $permissions = $DefaultPermissions['User']
            try {
                if ($obj.AdminCount -eq 1) {
                    $permissions = $DefaultPermissions['AdminSDHolder']
                }
            } catch {
            }
        }
        foreach ($p in $ps.keys) {
            if ([String]::IsNullOrEmpty($ps[$p])) {
                continue
            }
            $parameters = @{
                'XmlDoc' = $XmlDoc
                'NodeName' = $ps[$p]
                'Data' = $obj.$p
                'DefaultPermissions' = $permissions
                'Domain' = $Domain
            }
            if (![String]::IsNullOrEmpty($Server)) {
                $parameters.Add('Server', $Server)
            }
            if ($Credential) {
                $parameters.Add('Credential', $Credential)
            }
            $node = Create-SettingsXmlNode @parameters
            if ($node) {
                $x.AppendChild($node) | Out-Null
            }
            if ($node -eq $false) {
                return $false
            }
        }
    }
    return $true
}

function Backup-Computers {
<#
.SYNOPSIS
    Backs up Computer objects.
.DESCRIPTION
    Backs up Computer objects.
.PARAMETER XmlDoc
    Xml Document.
.PARAMETER XmlNode
    Xml Node.
.PARAMETER DefaultPermissions
    Default permissions on AD objects.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$XmlDoc = $null,

        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$XmlNode = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$DefaultPermissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    Add-Log -Log 'Backing up Computer objects' -Type Title2
    $ps = @{
        'AdminCount' = $null
        'Description' = $Script:DefDescription
        'DistinguishedName' = $Script:DefDistinguishedName
        'Name' = $Script:DefName
        'ProtectedFromAccidentalDeletion' = $Script:DefProtectedFromDel
        'SamAccountName' = $Script:DefSamAccountName
    }
    if ($DefaultPermissions) {
        $ps.Add('ntSecurityDescriptor', $Script:DefPermissions)
    }
    $defParams = @{
        'ResultSetSize' = $null
        'SearchScope' = [Microsoft.ActiveDirectory.Management.ADSearchScope]::Subtree
    }
    if (![String]::IsNullOrEmpty($Scope)) {
        $defParams.Add('SearchBase', $Scope)
    }
    $parameters = $defParams + @{
        'Properties' = @($ps.Keys)
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    $objs = $null
    $ousDN = $null
    try {
        $ousDN = Get-ADOrganizationalUnit -Filter '*' -Properties 'DistinguishedName' @defParams @commonParams
        $ousDN = @($ousDN.DistinguishedName)
        $objs = Get-ADComputer -Filter '*' @parameters @commonParams
        Add-Log -Log "Number of objects found: $(@($objs).Count)"
    } catch {
        Add-Log -Log "Error while listing Active Directory objects. $_" -Type Error
        return $false
    }
    $xs = $xmlDoc.CreateNode('element', $Script:DefComputers, $null)
    $xmlNode.AppendChild($xs) | Out-Null
    foreach ($obj in $objs) {
        if (!$obj) {
            continue
        }
        $path = $obj.DistinguishedName -ireplace '^.+?,(CN|OU.+)', '$1'
        if ($path -inotin $ousDN) {
            Add-Log -Log "Object '$($obj.DistinguishedName)' not in the scope (skipped)"
            continue
        }
        Add-Log -Log "Backing up object '$($obj.DistinguishedName)'"
        $x = $xmlDoc.CreateNode('element', $Script:DefComputer, $null)
        $xs.AppendChild($x) | Out-Null
        $permissions = $null
        if ($DefaultPermissions) {
            $permissions = $DefaultPermissions['Computer']
            try {
                if ($obj.AdminCount -eq 1) {
                    $permissions = $DefaultPermissions['AdminSDHolder']
                }
            } catch {
            }
        }
        foreach ($p in $ps.keys) {
            if ([String]::IsNullOrEmpty($ps[$p])) {
                continue
            }
            $parameters = @{
                'XmlDoc' = $XmlDoc
                'NodeName' = $ps[$p]
                'Data' = $obj.$p
                'DefaultPermissions' = $permissions
                'Domain' = $Domain
            }
            if (![String]::IsNullOrEmpty($Server)) {
                $parameters.Add('Server', $Server)
            }
            if ($Credential) {
                $parameters.Add('Credential', $Credential)
            }
            $node = Create-SettingsXmlNode @parameters
            if ($node) {
                $x.AppendChild($node) | Out-Null
            }
            if ($node -eq $false) {
                return $false
            }
        }
    }
    return $true
}

function Backup-Groups {
<#
.SYNOPSIS
    Backs up Group objects.
.DESCRIPTION
    Backs up Group objects.
.PARAMETER XmlDoc
    Xml Document.
.PARAMETER XmlNode
    Xml Node.
.PARAMETER DefaultPermissions
    Default permissions on AD objects.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$XmlDoc = $null,

        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$XmlNode = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$DefaultPermissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    Add-Log -Log 'Backing up Group objects' -Type Title2
    $ps = @{
        'AdminCount' = $null
        'Description' = $Script:DefDescription
        'DisplayName' = $Script:DefDisplayName
        'DistinguishedName' = $Script:DefDistinguishedName
        'GroupCategory' = $Script:DefCategory
        'GroupScope' = $Script:DefScope
        'Members' = $Script:DefMembers
        'Name' = $Script:DefName
        'ProtectedFromAccidentalDeletion' = $Script:DefProtectedFromDel
        'SamAccountName' = $Script:DefSamAccountName
    }
    if ($DefaultPermissions) {
        $ps.Add('ntSecurityDescriptor', $Script:DefPermissions)
    }
    $defParams = @{
        'ResultSetSize' = $null
        'SearchScope' = [Microsoft.ActiveDirectory.Management.ADSearchScope]::Subtree
    }
    if (![String]::IsNullOrEmpty($Scope)) {
        $defParams.Add('SearchBase', $Scope)
    }
    $parameters = $defParams + @{
        'Properties' = @($ps.Keys)
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    $objs = $null
    $ousDN = $null
    try {
        $ousDN = Get-ADOrganizationalUnit -Filter '*' -Properties 'DistinguishedName' @defParams @commonParams
        $ousDN = @($ousDN.DistinguishedName)
        $objs = Get-ADGroup -Filter '*' @parameters @commonParams
        Add-Log -Log "Number of objects found: $(@($objs).Count)"
    } catch {
        Add-Log -Log "Error while listing Active Directory objects. $_" -Type Error
        return $false
    }
    $xs = $xmlDoc.CreateNode('element', $Script:DefGroups, $null)
    $xmlNode.AppendChild($xs) | Out-Null
    foreach ($obj in $objs) {
        if (!$obj) {
            continue
        }
        $path = $obj.DistinguishedName -ireplace '^.+?,(CN|OU.+)', '$1'
        if ($path -inotin $ousDN) {
            Add-Log -Log "Object '$($obj.DistinguishedName)' not in the scope (skipped)"
            continue
        }
        Add-Log -Log "Backing up object '$($obj.DistinguishedName)'"
        $x = $xmlDoc.CreateNode('element', $Script:DefGroup, $null)
        $xs.AppendChild($x) | Out-Null
        $permissions = $null
        if ($DefaultPermissions) {
            $permissions = $DefaultPermissions['Group']
            try {
                if ($obj.AdminCount -eq 1) {
                    $permissions = $DefaultPermissions['AdminSDHolder']
                }
            } catch {
            }
        }
        $members = $null
        try {
            $members = Get-ADGroupMember -Identity $obj.DistinguishedName @commonParams
        } catch {
            Add-Log -Log "Unable to list group members. $_" -Type Error
            return $false
        }
        foreach ($p in $ps.keys) {
            if ([String]::IsNullOrEmpty($ps[$p])) {
                continue
            }
            $parameters = @{
                'XmlDoc' = $XmlDoc
                'NodeName' = $ps[$p]
                'Data' = $obj.$p
                'DefaultPermissions' = $permissions
                'Domain' = $Domain
            }
            if (![String]::IsNullOrEmpty($Server)) {
                $parameters.Add('Server', $Server)
            }
            if ($Credential) {
                $parameters.Add('Credential', $Credential)
            }
            if ($p -ieq 'Members') {
                $parameters['Data'] = $members
            }
            $node = Create-SettingsXmlNode @parameters
            if ($node) {
                $x.AppendChild($node) | Out-Null
            }
            if ($node -eq $false) {
                return $false
            }
        }
    }
    return $true
}

function Backup-WmiFilters {
<#
.SYNOPSIS
    Backs up Wmi Filter objects.
.DESCRIPTION
    Backs up Wmi Filter objects.
.PARAMETER XmlDoc
    Xml Document.
.PARAMETER XmlNode
    Xml Node.
.PARAMETER DefaultPermissions
    Default permissions on AD objects.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$XmlDoc = $null,

        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$XmlNode = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$DefaultPermissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    Add-Log -Log 'Backing up WMI Filter objects' -Type Title2
    $ps = @{
        'distinguishedName' = $null
        'msWMI-Name' = $Script:DefWmiName
        'msWMI-Parm1' = $Script:DefWmiParm1
        'msWMI-Parm2' = $Script:DefWmiParm2
        'msWMI-Author' = $Script:DefWmiAuthor
        'showInAdvancedViewOnly' = $Script:DefShowAdvView
    }
    $permissions = $null
    if ($DefaultPermissions) {
        $ps.Add('ntSecurityDescriptor', $Script:DefPermissions)
        $permissions = $DefaultPermissions['WmiFilter']
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    $wmiPath = $null
    $objs = $null
    try {
        $wmiPath = "CN=SOM,CN=WMIPolicy,$($Domain.SystemsContainer)"
        $objs = Get-ADObject -SearchBase $wmiPath `
                             -Filter { objectClass -eq 'msWMI-Som' } `
                             -Properties @($ps.Keys) `
                             @commonParams
        Add-Log -Log "Number of objects found: $(@($objs).Count)"
    } catch {
        Add-Log -Log "Error while listing existing WMI filters. $_" -Type Error
        return $false
    }
    $xs = $xmlDoc.CreateNode('element', $Script:DefWmiFilters, $null)
    $xmlNode.AppendChild($xs) | Out-Null
    foreach ($obj in $objs) {
        if (!$obj) {
            continue
        }
        Add-Log -Log "Backing up object '$($obj.DistinguishedName)'"
        $x = $xmlDoc.CreateNode('element', $Script:DefWmiFilter, $null)
        $xs.AppendChild($x) | Out-Null
        foreach ($p in $ps.keys) {
            if ([String]::IsNullOrEmpty($ps[$p])) {
                continue
            }
            $parameters = @{
                'XmlDoc' = $XmlDoc
                'NodeName' = $ps[$p]
                'Data' = $obj.$p
                'DefaultPermissions' = $permissions
                'Domain' = $Domain
            }
            if (![String]::IsNullOrEmpty($Server)) {
                $parameters.Add('Server', $Server)
            }
            if ($Credential) {
                $parameters.Add('Credential', $Credential)
            }
            $node = Create-SettingsXmlNode @parameters
            if ($node) {
                $x.AppendChild($node) | Out-Null
            }
            if ($node -eq $false) {
                return $false
            }
        }
    }
    return $true
}

function Backup-MigrationTable {
<#
.SYNOPSIS
    Builds a migration table after a backup of GPOs.
.DESCRIPTION
    Builds a migration table after a backup of GPOs.
.PARAMETER FilePath
    Save the migration table to this file.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER GpoBackupFolder
    Path to the folder where GPOs are backed-up.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    param (
        [parameter(mandatory=$true)]
        [String]$FilePath = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [parameter(mandatory=$true)]
        [String]$GpoBackupFolder = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    Add-Log -Log 'Building Migration Table'
    if ([String]::IsNullOrEmpty($FilePath) -or !(Test-Path -Path $FilePath -IsValid)) {
        Add-Log -Log 'Migration table file name must be a valid path' -Type Error
        return $false
    }
    if (!(Test-Path -Path $FilePath)) {
        try {
            New-Item -Path $FilePath -ItemType File -Force -ErrorAction Stop | Out-Null
            $FilePath = (Resolve-Path -Path $FilePath -ErrorAction Stop).Path
        } catch {
            Add-Log -Log "Unable to create the file '$FilePath'. $_" -Type Error
            return $false
        }
    }
    try {
        $FilePath = (Resolve-Path -Path $FilePath -ErrorAction Stop).Path
    } catch {
        Add-Log -Log "Unable to resolve path '$FilePath'. $_" -Type Error
        return $false
    }
    try {
        $GpoBackupFolder = (Resolve-Path -Path $GpoBackupFolder -ErrorAction Stop).Path
    } catch {
        Add-Log -Log "Unable to resolve path '$GpoBackupFolder'. $_" -Type Error
        return $false
    }
    $migtable = $null
    $success = $true
    try {
        $gpm = New-Object -ComObject gpmGMT.gpm
        $constants = $gpm.getConstants()
        $folder = $gpm.GetbackupDir($GpoBackupFolder)
        $crit = $gpm.CreateSearchCriteria()
        $gpos = $folder.SearchBackups($crit)
        $migtable = $gpm.CreateMigrationTable()
        foreach ($gpo in $gpos) {
            $migtable.Add($constants.ProcessSecurity, $gpo)
        }
        $migtable.Save($FilePath)
        $migtable = New-Object -TypeName Microsoft.GroupPolicy.GPMigrationTable($FilePath)
        foreach ($entry in $migtable.GetEntries()) {
            if (!$entry -or ($entry.EntryType -eq [Microsoft.GroupPolicy.GPEntryType]::UncPath)) {
                continue
            }
            $source = $entry.Source
            #$type = $entry.EntryType
            $dest = $entry.Destination
            $sam, $fqdn = $source -split '@'
            if ($sam -imatch '^S-\d+-\d+.*') {
                continue # Accounts with SID-like SamAccountName are not built-in accounts.
            }
            $sid = Convert-StringToSid -String $source -Server $Server -Credential $Credential -Domain $Domain
            if ($sid -eq $null) {
                if ([String]::IsNullOrEmpty($fqdn) -and ($sam -ieq 'Power Users')) {
                    $sid = [System.Security.Principal.SecurityIdentifier]'S-1-5-32-547' # Power Users
                    Add-Log -Log "Unknown entry '$source' will be translated to '$($sid)'" -Type Warning
                } else {
                    Add-Log -Log "Unknown entry '$source' found" -Type Warning
                    continue
                }
            }
            if ((Test-WellKnownSid -Sid $sid)) {
                $dest = Get-AnonymizedSid -Sid "$($sid)" -Domain $Domain
            }
            if (($dest -ieq '[DomainSID]-500') -and [String]::IsNullOrEmpty($fqdn)) {
                continue # Local Administrator account.
            }
            if ($dest -ine $entry.Destination) {
                $migtable.UpdateDestination($source, $dest) | Out-Null
            }
        }
        if ($success) {
            $migtable.Save($FilePath)
            Add-Log -Log "Migration table saved to '$FilePath'"
        }
    } catch {
        Add-Log -Log "Unable to build the migration table. $_" -Type Error
        $success = $false
    }
    $migtable = $null
    return $success
}

function Backup-Admx {
<#
.SYNOPSIS
    Backs up ADMX and ADML files from the Central Store.
.DESCRIPTION
    Backs up ADMX and ADML files from the Central Store.
.PARAMETER BackupFolder
    Backup folder.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [String]$BackupFolder = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    Add-Log -Log 'Backing up ADMX and ADML files from the Central Store' -Type Title2
    $dest = Join-Path -Path $BackupFolder -ChildPath $Script:AdmxFolderName
    if (!(Test-Path -Path $dest)) {
        try {
            $folder = New-Item -Path $dest -ItemType Directory -ErrorAction Stop
            Add-Log -Log "Created folder '$($folder.FullName)'"
        } catch {
            return $false
        }
    }
    try {
        $dest = (Resolve-Path -Path $dest -ErrorAction Stop).Path
    } catch {
        Add-Log -Log "Unable to resolve path backup folder path. $_" -Type Error
        return $false
    }
    if ([String]::IsNullOrEmpty($Server)) {
        $Server = $Domain.PDCEmulator
    }
    $parameters = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if ($Credential) {
        $parameters.Add('Credential', $Credential)
    }
    $store = "\\$($Server)\SYSVOL\$($Domain.DNSRoot)\Policies"
    $success = $true
    $drive = $null
    try {
        $drive = New-PSDrive -Name 'admx' -PSProvider FileSystem -Root $store @parameters
        Add-Log -Log "Mapped drive 'admx:\' to '$store'"
        $store = 'admx:\PolicyDefinitions'
        if (!(Test-Path -Path $store)) {
            Add-Log -Log "Path '$store' not found"
        } else {
            $items = Get-ChildItem -Path $store -Force -ErrorAction Stop
            foreach ($item in $items) {
                try {
                    Add-Log -Log "Copying item '$($item.FullName)' to '$dest'"
                    Copy-Item -Path $item.FullName `
                              -Destination $dest `
                              -Force `
                              -Recurse:$item.PSIsContainer `
                              -ErrorAction Stop
                } catch {
                    Add-Log -Log "Unable to copy item '$($item.Name)' to '$dest'" -Type Error
                    $success = $false
                    break
                }
            }
        }
    } catch {
        Add-Log -Log "Unable to copy items from '$store' to '$BackupFolder'. $_" -Type Error
        $success = $false
    } finally {
        try {
            $drive | Remove-PSDrive -Force -ErrorAction Stop
        } catch {
            Add-Log -Log "Unable to remove the PSDrive. $_" -Type Error
            $success = $false
        }
    }
    return $success
}

function Backup-GroupPolicies {
<#
.SYNOPSIS
    Backs up Group Policy objects.
.DESCRIPTION
    Backs up Group Policy objects.
.PARAMETER XmlDoc
    Xml Document.
.PARAMETER XmlNode
    Xml Node.
.PARAMETER All
    If true, all GPOs are backed up, not only linked GPOs.
.PARAMETER Reports
    If true, HTML reports are generated for each backed-up GPO.
.PARAMETER BackupFolder
    Backup folder.
.PARAMETER DefaultPermissions
    Default permissions on AD objects.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [System.Xml.XmlDocument]$XmlDoc = $null,

        [Parameter(Mandatory=$true)]
        [System.Xml.XmlElement]$XmlNode = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$All = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$Reports = $false,

        [Parameter(Mandatory=$true)]
        [String]$BackupFolder = $null,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$DefaultPermissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null
    )
    Add-Log -Log 'Backing up Group Policy objects' -Type Title2
    if (!(Test-Path -Path $BackupFolder)) {
        try {
            $folder = New-Item -Path $BackupFolder -ItemType Directory -ErrorAction Stop
            Add-Log -Log "Created folder '$($folder.FullName)'"
        } catch {
            return $false
        }
    }
    try {
        $BackupFolder = (Resolve-Path -Path $BackupFolder -ErrorAction Stop).Path
    } catch {
        Add-Log -Log "Unable to resolve path backup folder path. $_" -Type Error
        return $false
    }
    $defParams = @{
        'ResultSetSize' = $null
        'SearchScope' = [Microsoft.ActiveDirectory.Management.ADSearchScope]::Subtree
    }
    if (![String]::IsNullOrEmpty($Scope)) {
        $defParams.Add('SearchBase', $Scope)
    }
    $commonParams = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    if (![String]::IsNullOrEmpty($Server)) {
        $commonParams.Add('Server', $Server)
    }
    if ($Credential) {
        $commonParams.Add('Credential', $Credential)
    }
    $perms = $null
    if ($DefaultPermissions) {
        $perms = $DefaultPermissions['GroupPolicyObject']
    }
    $hGpos = @{}
    $ous = $null
    try {
        $ous = Get-ADOrganizationalUnit -Filter '*' -Properties 'DistinguishedName' @defParams @commonParams
        $ous = @($ous.DistinguishedName)
        if (![String]::IsNullOrEmpty($Scope) -and ($Scope -inotin $ous)) {
            $ous += $Scope
        }
        foreach ($ou in $ous) {
            if ([String]::IsNullOrEmpty($ou)) {
                continue
            }
            Add-Log -Log "Getting GPO Links from object '$ou'"
            $info = Get-XGPInheritance -Target $ou -Domain $Domain.DNSRoot @commonParams
            if (!$info) {
                continue
            }
            foreach ($link in $info.GpoLinks) {
                if (!$link -or !$link.GpoId) {
                    continue
                }
                $gpoId = $link.GpoId.ToString().ToLower()

                # GPOs 'Default Domain Policy' and 'Default Domain Controllers Policy' are not backed up.
                if (($gpoId -ieq '31b2f340-016d-11d2-945f-00c04fb984f9') -or
                    ($gpoId -ieq '6ac1786c-016f-11d2-945f-00c04fb984f9')) {
                    continue
                }
                $objLink = $link | Select-Object -Property GpoId, DisplayName, Enabled, Enforced, Target, Order
                $objLink.Target = $ou
                if ($hGpos.ContainsKey($gpoId)) {
                    $hGpos[$gpoId][$Script:DefLinks] += @($objLink)
                    continue
                }
                $gpo = Get-XGPO -Guid $gpoId -Domain $Domain.DNSRoot @commonParams
                $obj = Get-ADObject -Identity $gpo.Path -Properties 'ntSecurityDescriptor' @commonParams
                $data = @{
                    $Script:DefName = $gpo.DisplayName
                    $Script:DefTargetName = $gpo.DisplayName
                    $Script:DefGpoStatus = ($gpo.GpoStatus -as [Microsoft.GroupPolicy.GpoStatus])
                    #$Script:DefWmiFilter = $gpo.WmiFilter
                    $Script:DefLinks = @($objLink)
                }
                if ($perms) {
                    $data.Add($Script:DefPermissions, $obj.ntSecurityDescriptor)
                }
                $hGpos.Add($gpoId, $data)
            }
        }
    } catch {
        Add-Log -Log "Error while listing Active Directory objects. $_" -Type Error
        return $false
    }
    if ($All) {
        try {
            $gpos = Get-XGPO -All -Domain $Domain.DNSRoot @commonParams
            foreach ($gpo in $gpos) {
                $gpoId = $gpo.Id.ToString().ToLower()
                if ($hGpos.ContainsKey($gpoId)) {
                    continue
                }

                # GPOs 'Default Domain Policy' and 'Default Domain Controllers Policy' are not backed up.
                if (($gpoId -ieq '31b2f340-016d-11d2-945f-00c04fb984f9') -or
                    ($gpoId -ieq '6ac1786c-016f-11d2-945f-00c04fb984f9')) {
                    continue
                }
                $obj = Get-ADObject -Identity $gpo.Path -Properties 'ntSecurityDescriptor' @commonParams
                $data = @{
                    $Script:DefName = $gpo.DisplayName
                    $Script:DefTargetName = $gpo.DisplayName
                    $Script:DefGpoStatus = ($gpo.GpoStatus -as [Microsoft.GroupPolicy.GpoStatus])
                    #$Script:DefWmiFilter = $gpo.WmiFilter
                    $Script:DefLinks = $null
                }
                if ($perms) {
                    $data.Add($Script:DefPermissions, $obj.ntSecurityDescriptor)
                }
                $hGpos.Add($gpoId, $data)
            }
        } catch {
            Add-Log -Log "Error while listing Active Directory objects. $_" -Type Error
            return $false
        }
    }
    Add-Log -Log "Number of objects found: $($hGpos.Count)"
    $xs = $xmlDoc.CreateNode('element', $Script:DefGroupPolicies, $null)
    $xmlNode.AppendChild($xs) | Out-Null
    $migtableGpos = @()
    $padding = "$($hGpos.Count)".Length
    foreach ($gpoId in $hGpos.Keys) {
        $data = $hGpos[$gpoId]
        if (!$data) {
            continue
        }
        Add-Log -Log "Backing up object '$($data[$Script:DefName])'"
        $x = $xmlDoc.CreateNode('element', $Script:DefGroupPolicy, $null)
        $xs.AppendChild($x) | Out-Null
        foreach ($p in $data.Keys) {
            $parameters = @{
                'XmlDoc' = $XmlDoc
                'NodeName' = $p
                'Data' = $data[$p]
                'DefaultPermissions' = $perms
                'Domain' = $Domain
            }
            if (![String]::IsNullOrEmpty($Server)) {
                $parameters.Add('Server', $Server)
            }
            if ($Credential) {
                $parameters.Add('Credential', $Credential)
            }
            $node = Create-SettingsXmlNode @parameters
            if ($node) {
                $x.AppendChild($node) | Out-Null
            }
            if ($node -eq $false) {
                return $false
            }
        }
        try {
            $gpo = Backup-XGPO -Guid $gpoId -Path $BackupFolder -Domain $Domain.DNSRoot @commonParams
            if ($Reports) {
                $reportFolder = Join-Path -Path $BackupFolder -ChildPath $Script:GpoReportFolderName
                if (!(Test-Path -Path $reportFolder)) {
                    try {
                        $folder = New-Item -Path $reportFolder -ItemType Directory -ErrorAction Stop
                        Add-Log -Log "Created folder '$($folder.FullName)'"
                    } catch {
                        return $false
                    }
                }
                $reportAll = Join-Path -Path $reportFolder -ChildPath $Script:GpoAllReportsFolderName
                $reportDomain = Join-Path -Path $reportFolder -ChildPath $Script:GpoDomainReportsFolderName
                foreach ($f in @($reportAll, $reportDomain)) {
                    if (!(Test-Path -Path $f)) {
                        try {
                            $f = New-Item -Path $f -ItemType Directory -ErrorAction Stop
                            Add-Log -Log "Created folder '$($f.FullName)'"
                        } catch {
                            Add-Log -Log "Unable to cerate folder '$f'. $_" -Type Error
                            return $false
                        }
                    }
                }
                $name = $data[$Script:DefName]
                $ext = '.html'
                if ($name.StartsWith('*')) {
                    $name = $name.Replace('*', 'x') # For Tier Model GPOs, to have a 'clean' file name.
                }
                $maxSize = 260 - ($reportAll.Length + 2) - $ext.Length # FileNameMaxSize - baseFolder - FileExtension.
                $n = Get-SanitizedFileName -Name $name -MaxSize $maxSize
                if ($n -ine $name) {
                    $s = "Report for GPO '$($data[$Script:DefName])' ($($data[$Script:DefName].Length)) will be "
                    $s += "generated to '$($n)$($ext)' ($($n.Length + $ext.Length))"
                    Add-Log -Log $s -Type Warning
                }
                $name = $n
                $path = Join-Path -Path $reportAll -ChildPath "$($name)$($ext)"
                Get-XGPOReport -Guid $gpoId -ReportType Html -Path $path -Domain $Domain.DNSRoot @commonParams
                $n = $null
                try {
                    $links = Get-XmlValue -Xml $x -XmlPath $Script:DefLinks
                    if ($links -and (Test-Path -Path $path)) {
                        foreach ($link in $links.ChildNodes) {
                            $container = Get-XmlValue -Xml $link -XmlPath $Script:DefContainer
                            $order = Get-XmlValue -Xml $link -XmlPath $Script:DefOrder -DefaultValue 0
                            $enabled = Get-XmlValue -Xml $link -XmlPath $Script:DefEnabled -DefaultValue 'true'
                            $enforced = Get-XmlValue -Xml $link -XmlPath $Script:DefEnforced -DefaultValue 'false'
                            $order = "$order".PadLeft($padding, '0')
                            $folder = $reportDomain
                            if (![String]::IsNullOrEmpty($container)) {
                                $folder = $container -isplit ',?cn=|,?ou=|,?dc='
                                [Array]::Reverse($folder)
                                $maxSize = 248 - $reportDomain.Length # FolderNameMaxSize - baseFolder.
                                $illegal = $false
                                $p = ''
                                foreach ($subfolder in $folder) {
                                    if ([String]::IsNullOrEmpty($subfolder)) {
                                        continue
                                    }
                                    $maxSize -= $p.Length
                                    $f = Get-SanitizedFileName -Name $subfolder -MaxSize $maxSize
                                    if ($f -ine $subfolder) {
                                        $illegal = $true
                                    }
                                    $p += $f + '\'
                                }
                                $p = $p.Trim('\')
                                if ($illegal) {
                                    $s = "Converted path '$($container)' to '$($p)' to remove illegal characters"
                                    Add-Log -Log $s -Type Warning
                                }
                                $folder = Join-Path -Path $reportDomain -ChildPath $p
                            }
                            if (!(Test-Path -Path $folder)) {
                                New-Item -Path $folder -Force -ItemType Directory -ErrorAction Stop | Out-Null
                            }
                            if ($enabled -ieq 'true') {
                                $enabled = ''
                            } else {
                                $enabled = '-Disabled'
                            }
                            if ($enforced -ieq 'true') {
                                $enforced = '-Enforced'
                            } else {
                                $enforced = ''
                            }
                            $linkInfo = "[$($order)$($enabled)$($enforced)]"
                            $fullName = Join-Path -Path $folder -ChildPath "$($linkInfo)$($name)$($ext)"
                            if ($fullName.Length -ge 260) {
                                $maxSize = $fullName.Length - 260
                                if ($name.Length -ge $maxSize) {
                                    $n = Get-SanitizedFileName -Name $name -MaxSize ($name.Length - $maxSize - 1)
                                    $fullName = Join-Path -Path $folder -ChildPath "$($linkInfo)$($n)$($ext)"
                                    $s = "File '$path' ($($path.Length)) will be copied to "
                                    $s += "'$fullName' ($($fullName.Length))"
                                    Add-Log $s -Type Warning
                                }
                            }
                            Copy-Item -Path $path -Destination $fullName -Force -Confirm:$false -ErrorAction Stop
                        }
                    }
                } catch {
                    Add-Log -Log "Unable to copy file '$path' to '$n'. $_" -Type Warning
                }
            }
            $migtableGpos += $gpo
        } catch {
            Add-Log -Log "Error while backing up object '$($data[$Script:DefName])'. $_" -Type Error
            return $false
        }
    }
    $migtableFile = Join-Path -Path $BackupFolder -ChildPath $Script:MigrationTableFile
    if (!(Backup-MigrationTable -FilePath $migtableFile `
                                -GpoBackupFolder $BackupFolder `
                                -Domain $Domain `
                                -Server $Server `
                                -Credential $Credential)) {
        return $false
    }
    try {
        Get-ChildItem -Path $BackupFolder -File -Recurse -Force -ErrorAction Stop | Where-Object {
            (($_.Attributes -band [System.IO.FileAttributes]::Hidden) -eq [System.IO.FileAttributes]::Hidden)
        } | ForEach-Object {
            $_.Attributes = ($_.Attributes -bxor [System.IO.FileAttributes]::Hidden)
        }
    } catch {
        Add-Log -Log "Error while updating files' attributes in folder '$BackupFolder''. $_" -Type Error
        return $false
    }
    return $true
}

function Backup-Objects {
<#
.SYNOPSIS
    Backs up objects.
.DESCRIPTION
    Backs up objects.
.PARAMETER OutputFolder
    Collected data will be saved into this folder.
.PARAMETER OU
    If true, backs up Organizational Unit objects.
.PARAMETER User
    If true, backs up User objects.
.PARAMETER Computer
    If true, backs up Computer objects.
.PARAMETER Group
    If true, backs up Group objects.
.PARAMETER GPO
    Backs up Group Policy objects.
.PARAMETER WmiFilter
    If true, backs up Wmi Filter objects.
.PARAMETER Admx
    If true, backs up ADMX and ADML files from the Central Store.
.PARAMETER GpoReports
    If true, GPO Reports are generated.
.PARAMETER RedirectContainers
    If true, backs up redirection of containers.
.PARAMETER DefaultPermissions
    Active Directory objects default permissions.
.PARAMETER Domain
    Active Directory domain object.
.PARAMETER Scope
    DistinguishedName of an Active Directory container.
.PARAMETER Server
    Domain Controller to select for all queries.
.PARAMETER Credential
    Credential to use for Active Directory Authentication.
.PARAMETER Force
    Updates the object if it already exists.
.OUTPUTS
    True on success, false otherwise.
#>
    param(
        [Parameter(Mandatory=$true)]
        [String]$OutputFolder = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$OU = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$User = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$Computer = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$Group = $false,

        [Parameter(Mandatory=$false)]
        [ValidateSet('All', 'LinkedOnly', 'None')]
        [String]$GPO = 'None',

        [Parameter(Mandatory=$false)]
        [Switch]$WmiFilter = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$Admx = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$GpoReports = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$RedirectContainers = $false,

        [Parameter(Mandatory=$false)]
        [System.Collections.Hashtable]$DefaultPermissions = $null,

        [Parameter(Mandatory=$true)]
        [Microsoft.ActiveDirectory.Management.ADDomain]$Domain = $null,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [Switch]$Force = $false
    )

    # Prepares output folder.
    if ([String]::IsNullOrEmpty($OutputFolder) -or !(Test-Path -Path $OutputFolder -IsValid)) {
        Add-Log -Log 'OutputFolder must be a valid path' -Type Error
        return $false
    }
    if (!(Test-Path -Path $OutputFolder)) {
        try {
            New-Item -Path $OutputFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Add-Log -Log "Created folder '$OutputFolder'"
        } catch {
            Add-Log -Log "Unable to create folder '$OutputFolder'. $_" -Type Error
            return $false
        }
    }
    try {
        $test = Get-Item -Path $OutputFolder -ErrorAction Stop
        if ($test -isnot [System.IO.DirectoryInfo]) {
            Add-Log -Log "Item '$($OutputFolder)' must be a folder" -Type Error
            return $false
        }
    } catch {
        Add-Log -Log "Error while getting item '$($OutputFolder)'. $_" -Type Error
        return $false
    }
    try {
        $OutputFolder = (Resolve-Path -Path $OutputFolder -ErrorAction Stop).Path
    } catch {
        Add-Log -Log "Unable to resolve path '$OutputFolder'. $_" -Type Error
        return $false
    }
    $xmlFile = Join-Path -Path $OutputFolder -ChildPath "$($Script:SettingsXmlFile).xml"
    $timeStamp = $null
    if (!$Force) {
        $loop = 5
        while ((Test-Path -Path $xmlFile) -and ($loop-- -gt 0)) {
            Start-Sleep -Seconds 1
            $timeStamp = '-' + (Get-Date -Format 'yyyyMMddTHHmmss')
            $xmlFile = Join-Path -Path $OutputFolder -ChildPath "$($Script:SettingsXmlFile)$($timeStamp).xml"
        }
    }
    if ((Test-Path -Path $xmlFile)) {
        if (!$Force) {
            Add-Log -Log 'All valid Xml file names are already taken' -Type Error
            return $false
        } else {
            Add-Log -Log "File '$($xmlFile)' will be overwritten" -Type Warning
        }
    }
    $gpoBackupFolder = $null
    try {
        $gpoBackupFolder = Join-Path -Path $OutputFolder -ChildPath $Script:GpoBackupFolderName -ErrorAction Stop
    } catch {
        Add-Log -Log "Unable to build GPO Backup folder path. $_" -Type Error
        return $false
    }

    # Builds Xml file.
    $xmlDoc = New-Object -TypeName System.Xml.XmlDocument
    $node = $xmlDoc.CreateXmlDeclaration('1.0', 'utf-8', $null)
    $node = $xmlDoc.AppendChild($node)
    $x = $xmlDoc.CreateNode('element', $Script:DefConfiguration, $null)
    $xmlDoc.AppendChild($x) | Out-Null
    $xParams = $xmlDoc.CreateNode('element', $Script:DefParameters, $null)
    $x.AppendChild($xParams) | Out-Null
    $node = $xmlDoc.CreateNode('element', $Script:DefVersion, $null)
    $node.InnerText = $Script:ScriptVersion
    $xParams.AppendChild($node) | Out-Null
    if ($GPO) {
        $node = $xmlDoc.CreateNode('element', $Script:DefGpoBackupFolder, $null)
        $node.InnerText = ".\$($Script:GpoBackupFolderName)"
        $xParams.AppendChild($node) | Out-Null
    }
    $parameters = @{
        'XmlDoc' = $xmlDoc
        'XmlNode' = $x
        'DefaultPermissions' = $DefaultPermissions
        'Domain' = $Domain
        'Scope' = $Scope
        'Server' = $Server
        'Credential' = $Credential
    }
    $success = $true
    if ($OU) {
        $success = Backup-OrganizationalUnits -RedirectContainers:$RedirectContainers @parameters
        if (!$success) {
            return $false
        }
    }
    if ($User) {
        $success = Backup-Users @parameters
        if (!$success) {
            return $false
        }
    }
    if ($Computer) {
        $success = Backup-Computers @parameters
        if (!$success) {
            return $false
        }
    }
    if ($Group) {
        $success = Backup-Groups @parameters
        if (!$success) {
            return $false
        }
    }
    if ($WmiFilter) {
        $success = Backup-WmiFilters @parameters
        if (!$success) {
            return $false
        }
    }
    if ($Admx) {
        $success = Backup-Admx -BackupFolder $gpoBackupFolder -Domain $Domain -Server $Server -Credential $Credential
        if (!$success) {
            return $false
        }
    }
    if ($GPO -ine 'None') {
        $all = ($GPO -ieq 'All')
        $success = Backup-GroupPolicies -BackupFolder $gpoBackupFolder -All:$all -Reports:$GpoReports @parameters
        if (!$success) {
            return $false
        }
    }

    # Updates Schema Attributes.
    $success = Get-SchemaAttributesFromXml -Xml $xmlDoc -Server $Server -Credential $Credential
    if (!$success) {
        return $false
    }
    $success = Set-SchemaAttributesFromXml -Xml $xmlDoc
    if (!$success) {
        return $false
    }

    # Saves Xml file.
    try {
        $xmlDoc.Save($xmlFile)
        Add-Log -Log "Xml file saved to '$($xmlFile)'"
    } catch {
        Add-Log -Log "Unable to save file '$($xmlFile)'. $_" -Type Error
        return $false
    }
    return (Test-XmlSchema -File $xmlFile)
}

#endregion

#region Requirements

function Test-Requirements {
<#
.SYNOPSIS
    Checks script's requirements.
.DESCRIPTION
    Checks script's requirements.
    This function also sets the BaseFolder script's variable.
.PARAMETER Backup
    Is Backup mode enabled?
.PARAMETER Restore
    Is Restore mode enabled?
.PARAMETER OutputFolder
    Path to OutputFolder.
.PARAMETER SettingsXml
    Path to Settings.xml file.
.PARAMETER GPO
    True if GPOs must be backed up or restored.
.PARAMETER Permissions
    True if permissions must be backed up or restored.
.PARAMETER Scope
    Active Directory path to search.
.PARAMETER Server
    Domain controller to query.
.PARAMETER Credential
    Credential to use to query a domain controller.
.PARAMETER LogFile
    Path to the log file.
.OUTPUTS
    Custom object.
#>
    param (
        [Parameter(Mandatory=$false)]
        [Switch]$Backup = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$Restore = $false,

        [Parameter(Mandatory=$false)]
        [String]$OutputFolder = $null,

        [Parameter(Mandatory=$false)]
        [String]$SettingsXml = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$GPO = $false,

        [Parameter(Mandatory=$false)]
        [Switch]$Permissions = $false,

        [Parameter(Mandatory=$false)]
        [String]$Scope = $null,

        [Parameter(Mandatory=$false)]
        [String]$Server = $null,

        [Parameter(Mandatory=$false)]
        [Management.Automation.PSCredential]$Credential = $null,

        [Parameter(Mandatory=$false)]
        [String]$LogFile = $null
    )
    $ret = @{
        'Domain' = $null
        'RootDse' = $null
        'DefaultPermissions' = $null
        'Server' = $null
        'Scope' = $null
        'LogFile' = $null
        'Success' = $false
    }
    $parameters = @{
        'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
    }
    $success = $true

    # Administrative privileges.
    $isAdmin = Test-Admin
    $isDC = $false
    try {
        $isDC = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $isDC = ($isDC.ProductType -eq 2)
    } catch {
        Add-Log -Log "Unable to get operating system information. $_" -Type Error
        return $ret
    }
    if ($Restore -and $isDC -and !$isAdmin) {
        Add-Log -Log 'The script must be executed with Administrative privileges' -Type Error
        return $ret
    }

    # Script version.
    Add-Log -Log "Script version: $($Script:ScriptVersion)"

    # Windows version.
    if ((Get-CurrentWindowsVersion) -lt $Script:WindowsMinVersion) {
        Add-Log -Log "Windows $($Script:WindowsMinVersion) or more recent is required to run the script" `
                -Type Error
        return $ret
    }

    # PowerShell version.
    Add-Log -Log "PowerShell version: $($PSVersionTable.PSVersion)"
    if ($PSVersionTable.PSVersion -lt ([Version]'4.0')) {
        Add-Log -Log 'PowerShell 4 or more recent is required to run the script' -Type Error
        return $ret
    }

    # Creates temporary folder.
    try {
        for ($i = 0; $i -lt [Int]::MaxValue; $i++) {
            $tempPath = Join-Path -Path $env:TEMP -ChildPath "$i" -ErrorAction Stop
            if (!(Test-Path -Path $tempPath)) {
                New-Item -Path $tempPath -ItemType Directory -ErrorAction Stop -Force | Out-Null
                $Script:TempFolder = $tempPath
                Add-Log -Log "Created folder '$($Script:TempFolder)'"
                break
            }
        }
    } catch {
        Add-Log -Log "Unable to create folder in '$($env:TEMP)'. $_" -Type Error
        return $ret
    }
    if ([String]::IsNullOrEmpty($Script:TempFolder)) {
        Add-Log -Log "Unable to create folder in '$($env:TEMP)'" -Type Error
        return $ret
    }

    # Script's parameters.
    Add-Log -Log 'Checking script''s parameters'
    if ($Backup) {
        if ([String]::IsNullOrEmpty($OutputFolder)) {
            $OutputFolder = Join-Path -Path $env:TEMP -ChildPath 'Manage-AdObjects'
        }
        if (!(Test-Path -Path $OutputFolder -IsValid)) {
            Add-Log -Log "OutputFolder '$OutputFolder' is not a valid path" -Type Error
            return $ret
        }
        $created = $false
        if (!(Test-Path -Path $OutputFolder)) {
            try {
                New-Item -Path $OutputFolder -ItemType Directory -Force @parameters | Out-Null
                $created = $true
            } catch {
                Add-Log -Log "Unable to create folder '$OutputFolder'. $_" -Type Error
                return $ret
            }
        }
        try {
            $OutputFolder = (Resolve-Path -Path $OutputFolder @parameters).Path
        } catch {
            Add-Log -Log "Unable to resolve path '$OutputFolder'. $_" -Type Error
            return $ret
        }
        $Script:BaseFolder = $OutputFolder
        if ($created) {
            Add-Log -Log "Created output folder: '$OutputFolder'"
        } else {
            Add-Log -Log "Output folder: '$OutputFolder'"
        }
        $pathMaxSize = 247
        if ($GPO) {
            try {
                $gpmc = New-Object -ComObject 'GPMgmt.GPM' @parameters
                $gpmc = $null
            } catch {
                Add-Log -Log 'Group Policy Management console must be installed to back up GPOs' -Type Error
                return $ret
            }
            $pathMaxSize = 150
        }
        if ($OutputFolder.Length -gt $pathMaxSize) {
            Add-Log -Log "Folder name must not exceed $pathMaxSize characters (length is $($OutputFolder.Length))" `
                    -Type Error
            return $ret
        }
    }
    if ($Restore) {
        if ([String]::IsNullOrEmpty($SettingsXml) -or !(Test-Path -Path $SettingsXml -IsValid)) {
            Add-Log -Log 'SettingsXml must be a valid path' -Type Error
            return $ret
        }
        if (!(Test-Path -Path $SettingsXml)) {
            Add-Log -Log "Path '$SettingsXml' not found" -Type Error
            return $ret
        }
        try {
            $path = (Resolve-Path -Path $SettingsXml @parameters).Path
            $Script:BaseFolder = Split-Path -Path $path -Parent @parameters
            Add-Log -Log "Working folder set to '$($Script:BaseFolder)'"
        } catch {
            Add-Log -Log "Unable to resolve path '$SettingsXml'. $_" -Type Error
            return $ret
        }
    }
    if (![String]::IsNullOrEmpty($LogFile)) {
        if (!(Test-Path -Path $LogFile -IsValid)) {
            Add-Log -Log "LogFile '$LogFile' is not a valid path" -Type Error
            return $ret
        }
        $parent = Split-Path -Path $LogFile -Parent -ErrorAction SilentlyContinue
        if ([String]::IsNullOrEmpty($parent)) {
            $parent = $Script:BaseFolder
            $LogFile = Join-Path -Path $Script:BaseFolder -ChildPath $LogFile
        }
        if ((Test-Path -Path $LogFile) -and !$Force) {
            Add-Log -Log "Log file '$LogFile' already exists" -Type Error
            return $ret
        }
        $created = $false
        if (!(Test-Path -Path $LogFile)) {
            try {
                New-Item -Path $LogFile -ItemType File -Force @parameters | Out-Null
                $created = $true
            } catch {
                Add-Log -Log "Unable to create file '$parent'. $_" -Type Error
                return $ret
            }
        }
        try {
            $LogFile = (Resolve-Path -Path $LogFile @parameters).Path
            if ($created) {
                #Add-Log -Log "Created file '$LogFile'"
            }
        } catch {
            Add-Log -Log "Unable to resolve path '$LogFile'. $_" -Type Error
            return $ret
        }
        $ret['LogFile'] = $LogFile
        Add-Log -Log "Log file: $($LogFile)"
    }

    # PowerShell modules.
    Add-Log -Log 'Loading PowerShell modules'
    $psModules = @(
        'ActiveDirectory'
        'NetSecurity'
    )
    if ($GPO) {
        $psModules += 'GroupPolicy'
    }
    #$env:ADPS_LoadDefaultDrive = 0
    foreach ($psModule in $psModules) {
        try {
            Import-Module -Name $psModule -Scope 'Global' -Force @parameters | Out-Null
        } catch {
            Add-Log -Log "Unable to import PowerShell module '$psModule'. $_" -Type Error
            $success = $false
        }
    }
    if (!$success) {
        return $ret
    }

    # Active Directory connection.
    Add-Log -Log 'Checking Active Directory connection'
    if (![String]::IsNullOrEmpty($Server)) {
        $parameters.Add('Server', $Server)
    }
    if ($Credential) {
        $parameters.Add('Credential', $Credential)
        Add-Log -Log "Credential to be used: $($Credential.UserName)"
    } else {
        Add-Log -Log "Credential to be used: $($env:USERDOMAIN)\$($env:USERNAME)"
    }
    $domain = $null
    try {
        $domain = Get-ADDomain @parameters
        if (!$parameters.ContainsKey('Server')) {
            $Server = $domain.PDCEmulator
            $parameters.Add('Server', $Server)
        }
        Add-Log -Log "Target Active Directory domain: '$($domain.DNSRoot)'"
        Add-Log -Log "Target Domain Controller: '$Server'"
        $rootSid = $domain.DomainSID
        $rootNetbiosName = $domain.NetBIOSName
        $rootPDCEmulator = $domain.PDCEmulator
        $forest = Get-ADForest @parameters
        $schMaster = $forest.SchemaMaster
        if ($domain.DNSRoot -ine $domain.Forest) {
            Add-Log -Log "Collecting information from the Schema Master: '$schMaster'"
            $parameters['Server'] = $schMaster
            $rootDomain = Get-ADDomain @parameters
            $parameters['Server'] = $Server
            $rootSid = $rootDomain.DomainSID
            $rootNetbiosName = $rootDomain.NetBIOSName
            $rootPDCEmulator = $rootDomain.PDCEmulator
            Add-Log -Log "Target Domain Controller in the Root domain: '$rootPDCEmulator'"
        }
        $ps = @{
            'ErrorAction' = [System.Management.Automation.ActionPreference]::Stop
            'MemberType' = [System.Management.Automation.PSMemberTypes]::NoteProperty
            'Force' = $true
        }
        $domain | Add-Member -Name 'RootDomainSID' -Value $rootSid @ps
        $domain | Add-Member -Name 'RootNetBIOSName' -Value $rootNetbiosName @ps
        $domain | Add-Member -Name 'RootPDCEmulator' -Value $rootPDCEmulator @ps
        $domain | Add-Member -Name 'SchemaMaster' -Value $schMaster @ps
    } catch {
        Add-Log -Log "Unable to get Active Directory information. $_" -Type Error
        $domain = $null
    }
    if (!$domain) {
        return $ret
    }
    $ret['Server'] = $Server
    $ret['Domain'] = $domain
    $rootDse = $null
    try {
        $rootDse = Get-ADRootDSE @parameters
    } catch {
        Add-Log -Log "Unable to get Active Directory Root DSE. $_" -Type Error
    }
    if (!$rootDse) {
        return $ret
    }
    $ret['RootDse'] = $rootDse
    if (![String]::IsNullOrEmpty($Scope)) {
        try {
            $obj = Get-ADObject -Identity $Scope -Properties 'DistinguishedName', 'ObjectClass' @parameters
            if ($obj) {
                if ($obj.ObjectClass -iin ('organizationalUnit', 'domainDNS')) {
                    $ret['Scope'] = $obj.DistinguishedName
                } else {
                    Add-Log -Log "Object '$($obj.DistinguishedName)' is not a valid container" -Type Error
                    $success = $false
                }
            } else {
                Add-Log -Log "Active Directory object '$Scope' not found" -Type Error
                $success = $false
            }
        } catch {
            Add-Log -Log "Object '$Scope' not found. $_" -Type Error
            $success = $false
        }
    } else {
        $ret['Scope'] = $domain.DistinguishedName
    }
    if ($success) {
        Add-Log -Log "Active Directory scope: '$($ret['Scope'])'"
    } else {
        return $ret
    }

    # Gets Active Directory Schema attributes.
    if ($Restore -and $Permissions) {
        Add-Log -Log 'Getting Active Directory Schema information'
        $success = Get-SchemaAttributesFromXml -Path $SettingsXml @parameters
        if (!$success) {
            return $ret
        }
    }

    # Gets Active Directory objects default permissions.
    if ($Permissions) {
        Add-Log -Log 'Getting Active Directory objects default permissions'
        $defPerms = @{
            'Computer' = $null
            'Container' = $null
            'Group' = $null
            'Group-Policy-Container' = $null
            'ms-WMI-Som' = $null
            'Organizational-Person' = $null
            'Organizational-Unit' = $null
            'Person' = $null
            'Top' = $null
            'User' = $null
        }
        $cns = @() + $defPerms.Keys
        $ads = New-Object -TypeName System.DirectoryServices.ActiveDirectorySecurity
        foreach ($cn in $cns) {
            try {
                $dn = "CN=$($cn),$($rootDse.schemaNamingContext)"
                $o = Get-ADObject -Identity $dn -Properties 'defaultSecurityDescriptor' @parameters
                $ads.SetSecurityDescriptorSddlForm($o.defaultSecurityDescriptor)
                $defPerms[$cn] = $ads.Access
            } catch {
                Add-Log -Log "Unable to get Schema default permissions for '$cn' class. $_" -Type Error
                return $ret
            }
        }
        try {
            $dn = "CN=AdminSDHolder,$($domain.SystemsContainer)"
            $o = Get-ADObject -Identity $dn -Properties 'ntSecurityDescriptor' @parameters
            $defPerms.Add('AdminSDHolder', $o.ntSecurityDescriptor.Access)
        } catch {
            Add-Log -Log "Unable to get AdminSDHolder permissions. $_" -Type Error
            return $ret
        }
        $ads = New-Object -TypeName System.DirectoryServices.ActiveDirectorySecurity
        $rights = [System.DirectoryServices.ActiveDirectoryRights]::CreateChild -bor
                  [System.DirectoryServices.ActiveDirectoryRights]::DeleteChild -bor
                  [System.DirectoryServices.ActiveDirectoryRights]::Self -bor
                  [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty -bor
                  [System.DirectoryServices.ActiveDirectoryRights]::DeleteTree -bor
                  [System.DirectoryServices.ActiveDirectoryRights]::Delete -bor
                  [System.DirectoryServices.ActiveDirectoryRights]::GenericRead -bor
                  [System.DirectoryServices.ActiveDirectoryRights]::WriteDacl -bor
                  [System.DirectoryServices.ActiveDirectoryRights]::WriteOwner
        $ace = New-Object -TypeName System.DirectoryServices.ActiveDirectoryAccessRule(
                [System.Security.Principal.SecurityIdentifier]'S-1-3-0', # Creator Owner.
                $rights,
                [System.Security.AccessControl.AccessControlType]::Allow,
                [System.DirectoryServices.ActiveDirectorySecurityInheritance]::Descendents
        )
        $ads.AddAccessRule($ace) # Default permissions inherited from SYSVOL.
        $defPerms['Group-Policy-Container'] += $ads.Access
        $ret['DefaultPermissions'] = @{
            'OrganizationalUnit' = @(
                $defPerms['Top']
                $defPerms['Organizational-Unit']
            )
            'User' = @(
                $defPerms['Top']
                $defPerms['Person']
                $defPerms['Organizational-Person']
                $defPerms['User']
            )
            'Computer' = @(
                $defPerms['Top']
                $defPerms['Person']
                $defPerms['Organizational-Person']
                $defPerms['User']
                $defPerms['Computer']
            )
            'Group' = @(
                $defPerms['Top']
                $defPerms['Group']
            )
            'GroupPolicyObject' = @(
                $defPerms['Top']
                $defPerms['Container']
                $defPerms['Group-Policy-Container']
            )
            'WmiFilter' = @(
                $defPerms['Top']
                $defPerms['ms-WMI-Som']
            )
            'AdminSDHolder' = @(
                $defPerms['AdminSDHolder']
            )
        }
    }
    $ret['Success'] = $success
    return $ret
}

#endregion

#region Script's entry point

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$success = $true
$req = $null
do {
    $parameters = @{
        'OU' = $OU
        'User' = $User
        'Computer' = $Computer
        'Group' = $Group
        'GPO' = $GPO
        'RedirectContainers' = $RedirectContainers
        'WMI' = $WmiFilter
        'Admx' = $Admx
    }
    $commonParams = @{
        'Scope' = $Scope
        'Server' = $Server
        'Credential' = $Credential
    }

    # Checks requirements.
    Add-Log 'Checking requirements' -Type Title1
    $req = Test-Requirements -Backup:$Backup `
                             -Restore:$Restore `
                             -OutputFolder $OutputFolder `
                             -SettingsXml $SettingsXml `
                             -GPO:($GPO -ine 'None') `
                             -Permissions:$Permissions `
                             -LogFile $LogFile `
                             @commonParams
    if ($PSBoundParameters.ContainsKey('GpoLinks') -and !$PSBoundParameters.ContainsKey('GPO')) {
        Add-Log -Log 'Parameter -GpoLinks is ignored because parameter -GPO was not specified' -Type Warning
    }
    if (!$req -or !$req.Success) {
        $success = $false
        break
    }
    $commonParams['Scope'] = $req['Scope']
    $commonParams['Server'] = $req['Server']
    if ($Confirm -and !(Get-Confirmation)) {
        $success = $false
        break
    }

    # Restores Active Directory objects.
    if ($Restore) {
        Add-Log 'Restoring Active Directory objects' -Type Title1
        $xml = Read-XmlSettings -Path $SettingsXml
        if (!$xml) {
            $success = $false
            break
        }
        $success = Restore-Objects -Xml $xml `
                                   -GpoLinks $GpoLinks `
                                   -Force:$Force `
                                   -Domain $req['Domain'] `
                                   -Permissions:$Permissions `
                                   @parameters `
                                   @commonParams
        if (!$success) {
            break
        }
    }

    # Backs up Active Directory objects.
    if ($Backup) {
        Add-Log 'Backing up Active Directory objects' -Type Title1
        $perms = $null
        if ($Permissions) {
            $perms = $req['DefaultPermissions']
        }
        $success = Backup-Objects -OutputFolder $Script:BaseFolder `
                                  -Force:$Force `
                                  -Domain $req['Domain'] `
                                  -DefaultPermissions $perms `
                                  -GpoReports:$GpoReports `
                                  @parameters `
                                  @commonParams
        if (!$success) {
            break
        }
    }
} while ($false)

# Removes temporary files.
if (![String]::IsNullOrEmpty($Script:TempFolder) -and (Test-Path -Path $Script:TempFolder)) {
    Remove-Item -Path $Script:TempFolder -Recurse -Force -ErrorAction SilentlyContinue
    Add-Log -Log "Removed temporary folder '$($Script:TempFolder)'"
}

# Builds result.
if (![String]::IsNullOrEmpty($req['LogFile'])) {
    if (!(Get-LogFile -File $req['LogFile'] -Format $LogFormat -Force)) {
        $success = $false
    }
}
if ($sw) {
    $sw.Stop()
    $executionTime = $sw.Elapsed.ToString('hh\:mm\:ss')
    Add-Log -Log "Script's execution time (hour:min:sec): $($executionTime)"
}
if ($success) {
    Add-Log -Log 'Script executed successfully' -Type Success
} else {
    Add-Log -Log 'Failed to execute the script' -Type Error
}
return $success

#endregion

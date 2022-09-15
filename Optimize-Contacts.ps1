
<#PSScriptInfo

.VERSION 1.0

.GUID 450baa1a-8f27-4137-a2ec-fa231d5e5862

.AUTHOR Tigran TIKSN Torosyan

.COMPANYNAME TIKSN Lab

.COPYRIGHT Tigran TIKSN Torosyan

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

#Requires -Module Microsoft.Graph

<# 

.DESCRIPTION 
 Remove Duplicate Contacts 

#> 
[CmdletBinding(
    SupportsShouldProcess = $true,
    ConfirmImpact = 'Low')]
param (
)

try {
    Import-Module -Name Microsoft.Graph

    Connect-Graph -Scopes @('User.Read', 'Contacts.Read', 'Contacts.ReadWrite')
    $mgUser = Get-MgUser
    Write-Information "Microsoft Graph user is $($mgUser.DisplayName)"

    # $userContacts = Get-MgUserContact -UserId $mgUser.Id -All
    $userContacts = Get-MgUserContact -UserId $mgUser.Id -Skip 0 -Top 1000

    $contactGroups = $userContacts
    | Group-Object -Property BusinessPhone , MobilePhone
    | Where-Object { $PSItem.Count -gt 1 }

    foreach ($contactGroup in $contactGroups) {
        $firstContactInGroup = $contactGroup.Group[0]
        for ($i = 1; $i -lt $contactGroup.Group.Count; $i++) {
            $anotherContactInGroup = $contactGroup.Group[$i]

            $areIdentical = (
                ($firstContactInGroup.AssistantName -eq $anotherContactInGroup.AssistantName) -and
                ($firstContactInGroup.Birthday -eq $anotherContactInGroup.Birthday) -and
                ($firstContactInGroup.BusinessAddress -eq $anotherContactInGroup.BusinessAddress) -and
                ($firstContactInGroup.BusinessHomePage -eq $anotherContactInGroup.BusinessHomePage) -and
                ($firstContactInGroup.BusinessPhones -eq $anotherContactInGroup.BusinessPhones) -and
                ($firstContactInGroup.CompanyName -eq $anotherContactInGroup.CompanyName) -and
                ($firstContactInGroup.DisplayName -eq $anotherContactInGroup.DisplayName) -and
                ($firstContactInGroup.EmailAddresses -eq $anotherContactInGroup.EmailAddresses) -and
                ($firstContactInGroup.Extensions -eq $anotherContactInGroup.Extensions) -and
                ($firstContactInGroup.FileAs -eq $anotherContactInGroup.FileAs) -and
                ($firstContactInGroup.Generation -eq $anotherContactInGroup.Generation) -and
                ($firstContactInGroup.GivenName -eq $anotherContactInGroup.GivenName) -and
                ($firstContactInGroup.HomeAddress -eq $anotherContactInGroup.HomeAddress) -and
                ($firstContactInGroup.HomePhones -eq $anotherContactInGroup.HomePhones) -and
                ($firstContactInGroup.ImAddresses -eq $anotherContactInGroup.ImAddresses) -and
                ($firstContactInGroup.Initials -eq $anotherContactInGroup.Initials) -and
                ($firstContactInGroup.JobTitle -eq $anotherContactInGroup.JobTitle) -and
                ($firstContactInGroup.Manager -eq $anotherContactInGroup.Manager) -and
                ($firstContactInGroup.MiddleName -eq $anotherContactInGroup.MiddleName) -and
                ($firstContactInGroup.MobilePhone -eq $anotherContactInGroup.MobilePhone) -and
                ($firstContactInGroup.MultiValueExtendedProperties -eq $anotherContactInGroup.MultiValueExtendedProperties) -and
                ($firstContactInGroup.NickName -eq $anotherContactInGroup.NickName) -and
                ($firstContactInGroup.OfficeLocation -eq $anotherContactInGroup.OfficeLocation) -and
                ($firstContactInGroup.OtherAddress -eq $anotherContactInGroup.OtherAddress) -and
                ($firstContactInGroup.PersonalNotes -eq $anotherContactInGroup.PersonalNotes) -and
                ($firstContactInGroup.Photo -eq $anotherContactInGroup.Photo) -and
                ($firstContactInGroup.Profession -eq $anotherContactInGroup.Profession) -and
                ($firstContactInGroup.SingleValueExtendedProperties -eq $anotherContactInGroup.SingleValueExtendedProperties) -and
                ($firstContactInGroup.SpouseName -eq $anotherContactInGroup.SpouseName) -and
                ($firstContactInGroup.Surname -eq $anotherContactInGroup.Surname) -and
                ($firstContactInGroup.Title -eq $anotherContactInGroup.Title) -and
                ($firstContactInGroup.YomiCompanyName -eq $anotherContactInGroup.YomiCompanyName) -and
                ($firstContactInGroup.YomiGivenName -eq $anotherContactInGroup.YomiGivenName) -and
                ($firstContactInGroup.YomiSurname -eq $anotherContactInGroup.YomiSurname) -and
                ($firstContactInGroup.AdditionalProperties -eq $anotherContactInGroup.AdditionalProperties)
            )
            $areIdentical
        }
    }
}
catch {
    throw $_
}
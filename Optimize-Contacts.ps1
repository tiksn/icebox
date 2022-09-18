
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

function ArePhysicalAddressesIdentical {
    param ($x, $y)

    if (($null -eq $x) -and ($null -eq $y)) {
        return $true
    }

    return (($x.City -eq $y.City) -and
    ($x.CountryOrRegion -eq $y.CountryOrRegion) -and
    ($x.PostalCode -eq $y.PostalCode) -and
    ($x.State -eq $y.State) -and
    ($x.Street -eq $y.Street))
}

function AreBusinessHomePagesIdentical {
    param ($x, $y)

    if (($null -eq $x) -and ($null -eq $y)) {
        return $true
    }

    if ($x -eq $y) {
        return $true
    }

    return $true
    # $xr = Invoke-WebRequest -Uri $x
    # $yr = Invoke-WebRequest -Uri $y
    # $xrs = $xr.StatusCode
    # $yrs = $yr.StatusCode

    # return $false
}

function AreObjectsIdentical {
    param ($x, $y)

    if (($null -eq $x) -and ($null -eq $y)) {
        return $true
    }

    $d = Compare-Object -ReferenceObject $x -DifferenceObject $y
    if ($d) {
        return $false
    }
    else {
        return $true
    }
}

function AreAdditionalPropertiesIdentical {
    param ($x, $y)

    if (($null -eq $x) -and ($null -eq $y)) {
        return $true
    }

    $d = Compare-Object -ReferenceObject $x -DifferenceObject $y
    if ($d) {
        return $false
    }
    else {
        return $true
    }
}

function ArePhotosIdentical {
    param ($x, $y)

    if (($null -eq $x) -and ($null -eq $y)) {
        return $true
    }

    if ($x.Id -eq $y.Id) {
        return $true
    }

    return (($x.Height -eq $y.Height) -and
    ($x.Width -eq $y.Width))
}

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
        Write-Debug -Message "Examining Group $($contactGroup.Name)"
        Write-Debug -Message "First Contact ID: $($firstContactInGroup.Id)"
        for ($i = 1; $i -lt $contactGroup.Group.Count; $i++) {
            $anotherContactInGroup = $contactGroup.Group[$i]

            Write-Debug -Message "Another Contact ID: $($anotherContactInGroup.Id)"

            $hasAllInfo = (
                ($firstContactInGroup.AssistantName -eq $anotherContactInGroup.AssistantName) -and
                ($firstContactInGroup.Birthday -eq $anotherContactInGroup.Birthday) -and
                (ArePhysicalAddressesIdentical $firstContactInGroup.BusinessAddress $anotherContactInGroup.BusinessAddress) -and
                (AreBusinessHomePagesIdentical $firstContactInGroup.BusinessHomePage $anotherContactInGroup.BusinessHomePage) -and
                (AreObjectsIdentical $firstContactInGroup.BusinessPhones $anotherContactInGroup.BusinessPhones) -and
                ($firstContactInGroup.CompanyName -eq $anotherContactInGroup.CompanyName) -and
                ($firstContactInGroup.DisplayName -eq $anotherContactInGroup.DisplayName) -and
                (AreObjectsIdentical $firstContactInGroup.EmailAddresses $anotherContactInGroup.EmailAddresses) -and
                ($firstContactInGroup.Extensions -eq $anotherContactInGroup.Extensions) -and
                ($firstContactInGroup.FileAs -eq $anotherContactInGroup.FileAs) -and
                ($firstContactInGroup.Generation -eq $anotherContactInGroup.Generation) -and
                ($firstContactInGroup.GivenName -eq $anotherContactInGroup.GivenName) -and
                (ArePhysicalAddressesIdentical $firstContactInGroup.HomeAddress $anotherContactInGroup.HomeAddress) -and
                (AreObjectsIdentical $firstContactInGroup.HomePhones $anotherContactInGroup.HomePhones) -and
                (AreObjectsIdentical $firstContactInGroup.ImAddresses $anotherContactInGroup.ImAddresses) -and
                ($firstContactInGroup.Initials -eq $anotherContactInGroup.Initials) -and
                ($firstContactInGroup.JobTitle -eq $anotherContactInGroup.JobTitle) -and
                ($firstContactInGroup.Manager -eq $anotherContactInGroup.Manager) -and
                ($firstContactInGroup.MiddleName -eq $anotherContactInGroup.MiddleName) -and
                ($firstContactInGroup.MobilePhone -eq $anotherContactInGroup.MobilePhone) -and
                ($firstContactInGroup.MultiValueExtendedProperties -eq $anotherContactInGroup.MultiValueExtendedProperties) -and
                ($firstContactInGroup.NickName -eq $anotherContactInGroup.NickName) -and
                ($firstContactInGroup.OfficeLocation -eq $anotherContactInGroup.OfficeLocation) -and
                (ArePhysicalAddressesIdentical $firstContactInGroup.OtherAddress $anotherContactInGroup.OtherAddress) -and
                (-not $anotherContactInGroup.PersonalNotes) -and
                (ArePhotosIdentical $firstContactInGroup.Photo $anotherContactInGroup.Photo) -and
                ($firstContactInGroup.Profession -eq $anotherContactInGroup.Profession) -and
                ($firstContactInGroup.SingleValueExtendedProperties -eq $anotherContactInGroup.SingleValueExtendedProperties) -and
                ($firstContactInGroup.SpouseName -eq $anotherContactInGroup.SpouseName) -and
                ($firstContactInGroup.Surname -eq $anotherContactInGroup.Surname) -and
                ($firstContactInGroup.Title -eq $anotherContactInGroup.Title) -and
                ($firstContactInGroup.YomiCompanyName -eq $anotherContactInGroup.YomiCompanyName) -and
                ($firstContactInGroup.YomiGivenName -eq $anotherContactInGroup.YomiGivenName) -and
                ($firstContactInGroup.YomiSurname -eq $anotherContactInGroup.YomiSurname) -and
                (AreAdditionalPropertiesIdentical $firstContactInGroup.AdditionalProperties $anotherContactInGroup.AdditionalProperties)
            )

            Write-Debug -Message "First Conact Has All Info: $hasAllInfo"

            $anotherContactHasAdditionalInfo = (-not $anotherContactInGroup.PersonalNotes)

            Write-Debug -Message "Another Contact Has Additional Info: $anotherContactHasAdditionalInfo"
        }
    }
}
catch {
    throw $_
}

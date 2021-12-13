<#
    .SYNOPSIS
    Clear-AutoComplete
   
    Michel de Rooij
    michel@eightwone.com
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.4, August 17th, 2021
    
    .DESCRIPTION
    This script allows you to clear one or more locations where recipient information 
    is cached, as this could influence end user experience with certain migration scenarios.
    You have the option to clear the AutoComplete stream (Name Cache) for Outlook and 
    OWA, the Suggested Contacts or the Recipient Cache (Exchange 2013 only).
    Note that Outlook in cached mode will cache the AutoComplete stream as well (in the OST).
    Option then is to run Outlook with the /cleanautocompletecache switch, or clear and 
    let Outlook recreate the OST.
	
    .LINK
    http://eightwone.com
    
    .NOTES
    Requires Microsoft Exchange Web Services (EWS) Managed API 1.2 or up
    and Exchange 2010 or up or Exchange Online.

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial release
    1.01    Added X-AnchorMailbox usage for impersonation
            Renamed parameter Mailbox to Identity
    1.1     Bug fix in clearing Suggested Contacts and RecipientCache
    1.2     Reverified and updated to fix minor issues
            Changed deletes to HardDelete
    1.21    Added WhatIf/Confirm support
            Added success operations to Verbose output
    1.3     Added Pattern parameter
            Code rewrite
    1.4     Added code to support OAuth in addition to Basic Auth
            Added code to support pattern with AutoComplete
            Added code to support processing one or more Identity 
            Added TrustAll parameter
    
    .PARAMETER Identity
    Identity of the one or more mailboxes to process

    .PARAMETER Server
    Exchange Client Access Server to use for Exchange Web Services. When ommited, script will attempt to 
    use Autodiscover.

    .PARAMETER Credentials
    Specify credentials to use. When not specified, current credentials are used.
    Credentials can be set using $Credentials= Get-Credential
              
    .PARAMETER Type
    Determines what cached information to clear. Option are:
    - Outlook           : AutoComplete stream (also known as Nickname Cache), which is used by Outlook
    - OWA               : OWA nickname cache
    - SuggestedContacts : Automatically created contacts
    - RecipientCache    : Automatically created recipients (Only Exchange 2013)
    - All               : All of the above

    Default is Outlook,OWA

    .PARAMETER Pattern
    Specifies one or more patterns of the entries to remove from OWA, SuggestedContacts or RecipientCache. Does
    not work with Autocomplete stream (will only remove all entries). Patterns accept wildcards, e.g. to remove 
    all entries from the domain name contoso.com, use *@contoso.com. You can also use DN patterns, such 
    as '/o=ExchangeLabs/*'.

    .PARAMETER TenantId
    Specifies the identity of the Tenant.

    .PARAMETER ClientId
    Specifies the identity of the application configured in Azure Active Directory.

    .PARAMETER Credentials
    Specify credentials to use with Basic Authentication. Credentials can be set using $Credentials= Get-Credential
    This parameter is mutually exclusive with CertificateFile, CertificateThumbprint and Secret. 

    .PARAMETER CertificateThumbprint
    Specify the thumbprint of the certificate to use with OAuth authentication. The certificate needs
    to reside in the personal store. When using OAuth, providing TenantId and ClientId is mandatory.
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificateFile
    Specify the .pfx file containing the certificate to use with OAuth authentication. When a password is required,
    you will be prompted or you can provide it using CertificatePassword.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificatePassword
    Sets the password to use with the specified .pfx file. The provided password needs to be a secure string, 
    eg. -CertificatePassword (ConvertToSecureString -String 'P@ssword' -Force -AsPlainText)

    .PARAMETER Secret
    Specifies the client secret to use with OAuth authentication. The secret needs to be provided as a secure string.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and CertificateThumbprint. 

    .PARAMETER TrustAll
    Specifies if all certificates should be accepted, including self-signed certificates.

    .EXAMPLE
    Clear-AutoComplete.ps1 -Mailbox User1 -Type All -Verbose

    Removes all autocomplete information for mailbox User1.

    .EXAMPLE
    $Credentials= Get-Credential
    Clear-AutoComplete.ps1 -Identity olrik@office365tenant.com -Credentials $Credentials 
 
    Get credentials and removes Auto Complete information from olrik@office365tenant.com's mailbox.

    .EXAMPLE
    Import-CSV users.csv1 | Clear-AutoComplete.ps1 -Impersonation

    Uses a CSV file to removes AutoComplete information for a set of mailboxes, using impersonation. The CSV file should contain a column named Identity containing identities.
#>

[cmdletbinding(
    SupportsShouldProcess= $true,
    ConfirmImpact= 'High'
)]
param(
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'BasicAuth')] 
    [string[]]$Identity,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [string]$Server,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [switch]$Impersonation,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [ValidateSet("Outlook", "OWA", "SuggestedContacts", "RecipientCache", "All")]
    [array]$Type= @("Outlook","OWA"),
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [string[]]$Pattern,
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuth')] 
    [System.Management.Automation.PsCredential]$Credentials,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [System.Security.SecureString]$Secret,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [String]$CertificateThumbprint,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [String]$CertificateFile,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [System.Security.SecureString]$CertificatePassword,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [string]$TenantId,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [string]$ClientId
)
#Requires -Version 3.0

Begin {
    # Errors
    $ERR_EWSDLLNOTFOUND                      = 1000
    $ERR_EWSLOADING                          = 1001
    $ERR_MAILBOXNOTFOUND                     = 1002
    $ERR_AUTODISCOVERFAILED                  = 1003
    $ERR_CANTACCESSMAILBOXSTORE              = 1004
    $ERR_PROCESSINGMAILBOX                   = 1005
    $ERR_INVALIDCREDENTIALS= 1007
    $ERR_PROBLEMIMPORTINGCERT= 1008
    $ERR_CERTNOTFOUND= 1009
    
    Function Get-EmailAddress( $Identity) {
        $address= [regex]::Match([string]$Identity, ".*@.*\..*", "IgnoreCase")
        if( $address.Success ) {
            return $address.value.ToString()
        }
        Else {
            # Use local AD to look up e-mail address using $Identity as SamAccountName
            $ADSearch= New-Object DirectoryServices.DirectorySearcher( [ADSI]"")
            $ADSearch.Filter= "(|(cn=$Identity)(samAccountName=$Identity)(mail=$Identity))"
            $Result= $ADSearch.FindOne()
            If( $Result) {
                $objUser= $Result.getDirectoryEntry()
                return $objUser.mail.toString()
            }
            else {
                return $null
            }
        }
    }

    Function Import-ModuleDLL {
        param(
            [string]$Name,
            [string]$FileName,
            [string]$Package,
            [string]$ValidateObjName
        )

        $AbsoluteFileName= Join-Path -Path $PSScriptRoot -ChildPath $FileName
        If ( Test-Path $AbsoluteFileName) {
            # OK
        }
        Else {
            If( $Package) {
                If( Get-Command -Name Get-Package -ErrorAction SilentlyContinue) {
                    If( Get-Package -Name $Package -ErrorAction SilentlyContinue) {
                        $AbsoluteFileName= (Get-ChildItem -ErrorAction SilentlyContinue -Path (Split-Path -Parent (get-Package -Name $Package | Sort-Object -Property Version -Descending | Select-Object -First 1).Source) -Filter $FileName -Recurse).FullName
                    }
                }
            }
        }

        If( $absoluteFileName) {
            $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
            If( $ModLoaded) {
                Write-Verbose ('Module {0} v{1} already loaded' -f $ModLoaded.Name, $ModLoaded.Version)
            }
            Else {
                Write-Verbose ('Loading module {0}' -f $absoluteFileName)
                try {
                    Import-Module -Name $absoluteFileName -Global -Force
                    Start-Sleep 1
                }
                catch {
                    Write-Error ('Problem loading module {0}: {1}' -f $Name, $error[0])
                    Exit $ERR_DLLLOADING
                }
                $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
                If( $ModLoaded) {
                    Write-Verbose ('Module {0} v{1} loaded' -f $ModLoaded.Name, $ModLoaded.Version)
                }
                Try {
                    If( $validateObjName) {
                        $null= New-Object -TypeName $validateObjName
                    }
                }
                Catch {
                    Write-Error ('Problem initializing test-object from module {0}: {1}' -f $Name, $_.Exception.Message)
                    Exit $ERR_DLLLOADING
                }
            }
        }
        Else {
            Write-Verbose ('Required module {0} could not be located' -f $FileName)
            Exit $ERR_DLLNOTFOUND
        }
    }

    Function Set-SSLVerification {
        param(
            [switch]$Enable,
            [switch]$Disable
        )

        Add-Type -TypeDefinition  @"
            using System.Net.Security;
            using System.Security.Cryptography.X509Certificates;
            public static class TrustEverything
            {
                private static bool ValidationCallback(object sender, X509Certificate certificate, X509Chain chain,
                    SslPolicyErrors sslPolicyErrors) { return true; }
                public static void SetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback= ValidationCallback; }
                public static void UnsetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback= null; }
        }
"@
        If($Enable) {
            Write-Verbose ('Enabling SSL certificate verification')
            [TrustEverything]::UnsetCallback()
        }
        Else {
            Write-Verbose ('Disabling SSL certificate verification')
            [TrustEverything]::SetCallback()
        }
    }

    Function ConvertAutocompleteStream-ToObject {
        param(
            [Byte[]]$AutocompleteStream
        )
        $AutoCompleteEntries= [System.Collections.ArrayList]@()

        # First 4 bytes are metadata
        $MetadataStart= $AutocompleteStream[0..3]
        # Next 4 bytes are major version
        $MajorVersion= [System.BitConverter]::ToUint32( $AutocompleteStream, 4)
        # Next 4 bytes are minor version
        $MinorVersion= [System.BitConverter]::ToUint32( $AutocompleteStream, 8)
        # Last 12 bytes are metadata
        $MetadataEnd= $AutocompleteStream[ ($AutocompleteStream.Length-12) .. ($AutocompleteStream.Length-1)]

        $NumRows= [System.BitConverter]::ToUint32( $AutocompleteStream, 12)

        # Data starts here
        $i= 16

        For( $row=0; $row -lt $NumRows; $row++) {

            Write-Host ('Row: {0}' -f $row)
            $NumProps= [System.BitConverter]::ToUint32( $AutocompleteStream, $i)
            $i+= 4

            For( $prop=0; $prop -lt $NumProps; $prop++) {

                Write-Host ('Row {0}, Property {1}' -f $row, $prop)

                $PropertyTag= [System.BitConverter]::ToUint32( $AutocompleteStream, $i)
                $PropertyReserved= [System.BitConverter]::ToUint32( $AutocompleteStream, $i+4)
                $PropertyValueUnion= [System.BitConverter]::ToUint64( $AutocompleteStream, $i+8)
                $i+= 16

                Write-Host ('Row {0}, Property {1}, PropertyTag {2}' -f $row, $prop, $PropertyTag)

                Switch( $PropertyTag) {

                    0x101F { #PT_MV_UNICODE

                        #Array of PT_UNICODE
                        $NumPTUnicode= [System.BitConverter]::ToUint32( $AutocompleteStream, $i)
                        $i+= 4

                        $PropertyData= [System.Collections.ArrayList]@()
                        For( $s=0; $s -lt $NumPTUnicode; $s++) {
                            $len= [System.BitConverter]::ToUint32( $AutocompleteStream, $i)
                            $i+= 4
                            $null= $PropertyData.Add( [System.Text.Encoding]::Unicode.GetString( $AutocompleteStream[ ($i) .. ($i+ $len) ]))
                            $i+= len
                        }

                    }

                    0x101E { #PT_MV_STRING8

                        #Array of PT_STRING8
                        $NumPTString= [System.BitConverter]::ToUint32( $AutocompleteStream, $i)
                        $i+= 4

                        $PropertyData= [System.Collections.ArrayList]@()
                        For( $s=0; $s -lt $NumPTString; $s++) {
                            $len= [System.BitConverter]::ToUint32( $AutocompleteStream, $i)
                            $i+= 4
                            $null= $PropertyData.Add( [System.Text.Encoding]::ASCII.GetString( $AutocompleteStream[ ($i) .. ($i+ $len) ]))
                            $i+= len
                        }

                    }

                    0x1102 { #PT_MV_BINARY
                        # Array of PT_BINARY, shouldn't appear in AutocompleteStream
                    }

                    0x0102 { #PT_BINARY
                        $len= [System.BitConverter]::ToUint32( $AutocompleteStream, $i)
                        $i+= 4

                        $PropertyData= $AutocompleteStream[ ($i) .. ($i+ $len)]
                        $i+= len
                    }

                    0x0048 { #PT_CLSID
                        # Guid, shouldn't appear in AutocompleteStream
                        $i+= 16
                    }

                    0x0002 { #PT_UNICODE
                        $len= [System.BitConverter]::ToUint32( $AutocompleteStream, $i)
                        $i+= 4

                        $PropertyData= [System.Text.Encoding]::Unicode.GetString( $AutocompleteStream[ ($i) .. ($i+ $len) ])
                        $i+= len
                    }

                    0x001E { #PT_STRING8
                        $len= [System.BitConverter]::ToUint32( $AutocompleteStream, $i)
                        $i+= 4

                        $PropertyData= [System.Text.Encoding]::ASCII.GetString( $AutocompleteStream[ ($i) .. ($i+ $len) ])
                        $i+= len
                    }


                    0x0002 { #PT_I2 
                        # No data, data in union
                        $PropertyData= $null
                    }
                    0x0003 { #PT_LONG
                        # No data, data in union
                        $PropertyData= $null
                    }
                    0x0004 { #PT_R4
                        # No data, data in union
                        $PropertyData= $null
                    }
                    0x0005 { #PT_DOUBLE
                        # No data, data in union
                        $PropertyData= $null
                    }
                    0x000B { #PT_BOOLEAN
                        # No data, data in union
                        $PropertyData= $null
                    }
                    0x0040 { #PT_SYSTIME
                        # No data, data in union
                        $PropertyData= $null
                    }
                    0x0014 { #PT_I8
                        # No data, data in union
                        $PropertyData= $null
                    }
                    default {
                        Write-Error ('Unknown PropertyTag in AutocompleteStream object: {0}' -f $PropertyTag)
                        Return $null
                    }
                }

                Write-Host ('Data: {0}' -f $PropertyData)

            }

        }
    }

    Function Clear-AutoCompleteStream( $EwsService, $EmailAddress, $Pattern) {
        Write-Host ('Processing AutoComplete stream for {0}' -f $EmailAddress)
        $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::Inbox, $EmailAddress)  
        $InboxFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $FolderId)

        # Construct search filter for class & subject
        $ItemSearchFilterCollection= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, 'IPM.Configuration.Autocomplete')

        # Set the view/properties to return
        $ItemView= New-Object Microsoft.Exchange.WebServices.Data.ItemView( 1)
        $ItemView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet( [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        
        $PR_ROAMING_BINARYSTREAM = [Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition]::New( 0x7c09, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
        $ItemView.PropertySet.Add( $PR_ROAMING_BINARYSTREAM)   

        $ItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        $ItemSearchResults= $InboxFolder.FindItems( $ItemSearchFilterCollection, $ItemView)

        If( $ItemSearchResults.Items.Count -gt 0) {
            ForEach( $Item in $ItemSearchResults.Items) {

#                $AutocompleteStream = $null
#                $null= $Item.TryGetProperty( $PR_ROAMING_BINARYSTREAM, [ref]$AutocompleteStream)
#                If( $null -ne $AutocompleteStream) {
#                    $OriginalStream= ConvertAutocompleteStream-ToObject -AutocompleteStream $AutocompleteStream
#                }
#                Else {
#                    If ($pscmdlet.ShouldProcess( 'AutoComplete Stream', 'Cannot process AutoComplete configuration, just clear')) {
                    If ($pscmdlet.ShouldProcess( 'AutoComplete Stream', 'Clear')) {
                        $res= $Item.Delete( [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
                        Write-Host 'Cleared AutoComplete stream' -Foreground Green
                    }
#                }
            }
        }
        Else {
            Write-Host 'No AutoComplete Stream item found'
        }
    }

    Function Clear-OWAAutoComplete( $EwsService, $EmailAddress, $Pattern) {
        Write-Host ('Processing OWA AutoComplete for {0}' -f $EmailAddress)
        Try { 
            $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $EmailAddress)  
            $UserConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($EwsService, 'OWA.AutocompleteCache', $FolderId, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
        }
        Catch {
            Write-Host 'No OWA AutoComplete item found'
            $UserConfig= $false
        }
        If( $UserConfig) {
            If ($pscmdlet.ShouldProcess( 'OWA AutoComplete', 'Clear')) {

                If( $Pattern) {
                    $xmlDoc = New-Object System.Xml.XmlDocument
                    $xmlData = [System.Text.Encoding]::UTF8.GetString( $UserConfig.XmlData).Substring(1)
                    If( $xmlData) {
                        $xmlDoc.loadXml( $xmlData)
                        $nodes= $xmlDoc.SelectNodes("/AutoCompleteCache/entry")
                        ForEach( $node in $nodes) {
                            Write-Verbose ('Evaluating {0}' -f $node.smtpAddr)
                            $IsMatch= $false
                            ForEach( $ThisPattern in $Pattern) {
                                $IsMatch= $IsMatch -or ($node.smtpAddr -like $ThisPattern)
                            }
                            If( $IsMatch) {
                                Write-Host ('Removing {0}' -f $node.smtpAddr) -Foreground Green
                                $node.parentNode.removeChild( $node) | Out-Null
                            }
                            Else {
                                Write-Verbose ('Skipping {0}' -f $node.smtpAddr)
                           }
                        }
                        $newXmlData= [System.Text.Encoding]::UTF8.GetBytes( [System.Text.Encoding]::UTF8.GetString( $UserConfig.XmlData).Substring(0,1) + $XmlDoc.OuterXml)
                        $UserConfig.XmlData= $NewXmlData
                        Try {
                            $UserConfig.Update()      
                            Write-Verbose 'Updated OWA AutoComplete'
                        }
                        Catch {
                            Write-Error ('Problem updating OWA Autocomplete: {0}' -f $error[0])
                        }
                    }
                    Else {
                        Write-Warning ('Problem retrieving UserConfiguration')
                    }
                }
                Else {
                    # Zap all
                    Try {
                        $UserConfig.Delete()
                        $UserConfig.Update()      
                        Write-Verbose 'Cleared OWA AutoComplete'
                    }
                    Catch {
                        Write-Error ('Problem removing OWA Autocomplete Stream item: {0}' -f $error[0])
                    }
                }
            }
        }
    }

    Function Clear-SuggestedContacts( $EwsService, $EmailAddress, $Pattern) {
        Write-Host ('Processing SuggestedContacts for {0}' -f $EmailAddress)
        $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::MsgFolderRoot, $EmailAddress)  
        $FolderView= New-Object Microsoft.Exchange.WebServices.Data.FolderView( 1)
        $FolderView.Traversal= [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
        $FolderSearchFilter= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, 'Suggested Contacts')
        $FolderSearchResults= $EwsService.FindFolders( $FolderId, $FolderSearchFilter, $FolderView)
        If( $FolderSearchResults.Count -gt 0) {
            ForEach( $Folder in $FolderSearchResults) {
                If ($pscmdlet.ShouldProcess( 'SuggestedContacts', 'Clear')) {
                    If( $Pattern) {
                        $ItemView= New-Object Microsoft.Exchange.WebServices.Data.FolderView( 1000)
                        $Items = $EwsService.FindItems( $Folder.Id, $ItemView)
                        If( $Items) {
                            ForEach( $Item in $Items.Items) {
                                $IsMatch= $false
                                ForEach( $ThisPattern in $Pattern) {
                                    $IsMatch= $IsMatch -or ($Item.EmailAddress1 -like $ThisPattern)
                                }
                                If( $IsMatch) {
                                    Try {
                                        $Item.delete("HardDelete")
                                        Write-Host ('Removing {0}' -f $Item.EmailAddress1) -Foreground Green
                                     }
                                     Catch {
                                         Write-Error ('Problem removing item {0} from SuggestedContacts folder: {0}' -f $Item.EmailAddress1, $error[0])
                                     }
                                }
                                Else {
                                    Write-Verbose ('Skipping {0}' -f $Item.EmailAddress1)
                                }
                            }
                        }
                        Else {
                            Write-Host 'No entries found in SuggestedContacts'
                        }
                    }
                    Else {
                        # Zap all
                        Try {
                            $Folder.Empty( [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
                            Write-Verbose "Cleared 'Suggested Contacts'"
                        }
                        Catch {
                            Write-Error ('Problem removing SuggestedContacts folder: {0}' -f $error[0])
                        }
                    }
                }
            }
        }
        Else {
            Write-Host 'No Suggested Contacts folder found'
        }        
    }

    Function Clear-RecipientCache( $EwsService, $EmailAddress, $Pattern) {
        Write-Host ('Processing RecipientCache for {0}' -f $EmailAddress)
        Try {
            $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecipientCache, $EmailAddress)  
            $RecipientCacheFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $FolderId)
        }
        Catch {
            Write-Host 'No RecipientCache folder found'
        }
        If( $RecipientCacheFolder) {
            If ($pscmdlet.ShouldProcess( 'RecipientCache', 'Clear')) {
                If( $Pattern) {
                    $ItemView= New-Object Microsoft.Exchange.WebServices.Data.FolderView( 1000)
                    $Items = $EwsService.FindItems( $RecipientCacheFolder.Id, $ItemView)
                    If( $Items) {
                        ForEach( $Item in $Items.Items) {
                            $IsMatch= $false
                            ForEach( $ThisPattern in $Pattern) {
                                $IsMatch= $IsMatch -or ($Item.EmailAddress1 -like $ThisPattern)
                            }
                            If( $IsMatch) {
                                Write-Host ('Removing {0}' -f $Item.EmailAddress1) -Foreground Green
                                $Item.delete("HardDelete")
                            }
                            Else {
                                Write-Verbose ('Skipping {0}' -f $Item.EmailAddress1)
                            }
                        }
                    }
                    Else {
                        Write-Host 'No entries found in RecipientCache'
                    }
                }
                Else {
                    # Zap All
                    Try {
                        $RecipientCacheFolder.Empty( [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
                        Write-Verbose "Cleared RecipientCache"
                    }
                    Catch {
                        Write-Error ('Problem removing RecipientCache folder: {0}' -f $error[0])
                    }
                }
            }
        }
    }

    If( $Pattern) {
        Try {
            ForEach( $ThisPattern in $Pattern) {
                $res= 'test' -like $Pattern
            }
            Write-Verbose ('Pattern(s) specified: {0}' -f ($Pattern -join ','))
        }
        Catch {
            Throw( 'Provided pattern does not seem to be a valid expression')
        }
    }
    Else {
        # Not specified, so zap 'em all
    }

    Import-ModuleDLL -Name 'Microsoft.Exchange.WebServices' -FileName 'Microsoft.Exchange.WebServices.dll' -Package 'Exchange.WebServices.Managed.Api' -validateObjName 'Microsoft.Exchange.WebServices.Data.ExchangeVersion'
    
    # Load MSAL DLL when OAuth is to be used
    If($TenantId) {
        Import-ModuleDLL -Name 'Microsoft.Identity.Client' -FileName 'Microsoft.Identity.Client.dll' -Package 'Microsoft.Identity.Client' -validateObjName 'Microsoft.Identity.Client.TokenCache'
    }

    $ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    $EwsService= [Microsoft.Exchange.WebServices.Data.ExchangeService]::new( $ExchangeVersion)

    If( $Credentials) {
        try {
            Write-Verbose ('Using credentials {0}' -f $Credentials.UserName)
            $EwsService.Credentials= [System.Net.NetworkCredential]::new( $Credentials.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Credentials.Password )))
        }
        catch {
            Write-Error ('Invalid credentials provided: {0}' -f $_.Exception.Message)
            Exit $ERR_INVALIDCREDENTIALS
        }
    }
    Else {
        # Use OAuth (and impersonation/X-AnchorMailbox always set)
        $Impersonation= $true

        If( $CertificateThumbprint -or $CertificateFile) {
            If( $CertificateFile) {
                
                # Use certificate from file using absolute path to authenticate
                $CertificateFile= (Resolve-Path -Path $CertificateFile).Path
                
                Try {
                    If( $CertificatePassword) {
                        $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $CertificatePassword)))
                    }
                    Else {
                        $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile)
                    }
                }
                Catch {
                    Write-Error ('Problem importing PFX: {0}' -f $_.Exception.Message)
                    Exit $ERR_PROBLEMIMPORTINGCERT
                }
            }
            Else {
                # Use provided certificateThumbprint to retrieve certificate from My store, and authenticate with that
                $CertStore= [System.Security.Cryptography.X509Certificates.X509Store]::new( [Security.Cryptography.X509Certificates.StoreName]::My, [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
                $CertStore.Open( [System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly )
                $X509Certificate2= $CertStore.Certificates.Find( [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $CertificateThumbprint, $False) | Select-Object -First 1
                If(!( $X509Certificate2)) {
                    Write-Error ('Problem locating certificate in My store: {0}' -f $error[0])
                    Exit $ERR_CERTNOTFOUND
                }
            }
            Write-Verbose ('Will use certificate {0}, issued by {1} and expiring {2}' -f $X509Certificate2.Thumbprint, $X509Certificate2.Issuer, $X509Certificate2.NotAfter)
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithCertificate( $X509Certificate2).withTenantId( $TenantId).Build()
               
        }
        Else {
            # Use provided secret to authenticate
            Write-Verbose ('Will use provided secret to authenticate')
            $PlainSecret= [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Secret))
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithClientSecret( $PlainSecret).withTenantId( $TenantId).Build()
        }
        $Scopes= New-Object System.Collections.Generic.List[string]
        $Scopes.Add( 'https://outlook.office365.com/.default')
        Try {
            $Response=$App.AcquireTokenForClient( $Scopes).executeAsync()
            $Token= $Response.Result
            $EwsService.Credentials= [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$Token.AccessToken
            Write-Verbose ('Authentication token acquired')
        }
        Catch {
            Write-Error ('Problem acquiring token: {0}' -f $error[0])
            Exit $ERR_INVALIDCREDENTIALS
        }
    }

    If( $TrustAll) {
        Set-SSLVerification -Disable
    }

}

Process {
        

    ForEach( $ThisIdentity in $Identity) {

        $EmailAddress= get-EmailAddress $ThisIdentity

        If( !$EmailAddress) {
            Write-Error ('Specified mailbox {0} not found' -f $ThisIdentity)
            Exit $ERR_MAILBOXNOTFOUND
        }
        Write-Host ('Processing mailbox {0}' -f $EmailAddress)

        If( $Impersonation) {
            Write-Verbose ('Using {0} for impersonation' -f $EmailAddress)
            $EwsService.ImpersonatedUserId= [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new( [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
            $EwsService.HttpHeaders.Clear()
            $EwsService.HttpHeaders.Add( 'X-AnchorMailbox', $EmailAddress)
        }
            
        If ($Server) {
            $EwsUrl= 'https://{0}/EWS/Exchange.asmx' -f $Server
            Write-Verbose ('Using Exchange Web Services URL {0}' -f $EwsUrl)
            $EwsService.Url= $EwsUrl
        }
        Else {
            Write-Verbose ('Looking up EWS URL using Autodiscover for {0}' -f $EmailAddress)
            try {
                # Set script to terminate on all errors (autodiscover failure isn't) to make try/catch work
                $ErrorActionPreference= 'Stop'
                $EwsService.autodiscoverUrl( $EmailAddress, {$true})
            }
            catch {
                Write-Error ('Autodiscover failed: {0}' -f $_.Exception.Message)
                Exit $ERR_AUTODISCOVERFAILED
            }
            $ErrorActionPreference= 'Continue'
            Write-Verbose 'Using EWS endpoint {0}' -f $EwsService.Url
        } 

        try {
            $RootFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::MsgFolderRoot)
        }
        catch {
            Write-Error ('Cannot access primary mailbox of {0}: {1}' -f $EmailAddress, $_.Exception.Message)
            Exit $ERR_CANTACCESSMAILBOXSTORE
        }

        if( $null -ne $RootFolder) {

            If( $Type -contains 'All' -or $Type -contains 'Outlook') {
                Clear-AutoCompleteStream $EwsService $EmailAddress $Pattern
            }

            If( $Type -contains 'All' -or $Type -contains 'OWA') {
                Clear-OWAAutoComplete $EwsService $EmailAddress $Pattern
            }

            If( $Type -contains 'All' -or $Type -contains 'SuggestedContacts') {
                Clear-SuggestedContacts $EwsService $EmailAddress $Pattern
            }

            If( $Type -contains 'All' -or $Type -contains 'RecipientCache') {
                Clear-RecipientCache $EwsService $EmailAddress $Pattern
            }
        }
    }
}
End {
    If( $TrustAll) {
        Set-SSLVerification -Enable
    }
}

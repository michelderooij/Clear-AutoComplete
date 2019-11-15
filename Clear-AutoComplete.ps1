<#
    .SYNOPSIS
    Clear-AutoComplete
   
    Michel de Rooij
    michel@eightwone.com
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.3, November 15th, 2019
    
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
    
    .PARAMETER Identity
    Identity of the Mailbox to process

    .PARAMETER Server
    Exchange Client Access Server to use for Exchange Web Services. When ommited, script will attempt to 
    use Autodiscover.

    .PARAMETER Credentials
    Specify credentials to use. When not specified, current credentials are used.
    Credentials can be set using $Credentials= Get-Credential
              
    .PARAMETER Impersonation
    When specified, uses impersonation for mailbox access, otherwise current logged on user is used.
    For details on how to configure impersonation access for Exchange 2010 using RBAC, see this article:
    http://msdn.microsoft.com/en-us/library/exchange/bb204095(v=exchg.140).aspx
    For details on how to configure impersonation for Exchange 2007, see KB article:
    http://msdn.microsoft.com/en-us/library/exchange/bb204095%28v=exchg.80%29.aspx

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

    .EXAMPLE
    Clear-AutoComplete.ps1 -Mailbox User1 -Type All -Verbose

    Removes all autocomplete information for mailbox User1.

    .EXAMPLE
    $Credentials= Get-Credential
    Clear-AutoComplete.ps1 -Identity olrik@office365tenant.com -Credentials $Credentials 
 
    Get credentials and removes Auto Complete information from olrik@office365tenant.com's mailbox.

    .EXAMPLE
    Import-CSV users.csv1 | Clear-AutoComplete.ps1 -Impersonation

    Uses a CSV file to removes AutoComplete information for a set of mailboxes, using impersonation.
#>

[cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact="High")]
param(
    [parameter( Position=0, Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    [string]$Identity,
    [parameter( Mandatory=$false)]
    [string]$Server,
    [parameter( Mandatory=$false)]
    [switch]$Impersonation,
    [parameter( Mandatory= $false)] 
    [System.Management.Automation.PsCredential]$Credentials,
    [parameter( Mandatory= $false)]
    [ValidateSet("Outlook", "OWA", "SuggestedContacts", "RecipientCache", "All")]
    [array]$Type= @("Outlook","OWA"),
    [parameter( Mandatory= $false)]
    [string[]]$Pattern
)

process {
    # Errors
    $ERR_EWSDLLNOTFOUND                      = 1000
    $ERR_EWSLOADING                          = 1001
    $ERR_MAILBOXNOTFOUND                     = 1002
    $ERR_AUTODISCOVERFAILED                  = 1003
    $ERR_CANTACCESSMAILBOXSTORE              = 1004
    $ERR_PROCESSINGMAILBOX                   = 1005
    
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

    Function Load-EWSManagedAPIDLL {
        $EWSDLL= 'Microsoft.Exchange.WebServices.dll'
        If( Test-Path (Join-Path $pwd $EWSDLL)) {
            $EWSDLLPath= $pwd
        }
        Else {
            $EWSDLLPath = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory'))
            if (!( Test-Path (Join-Path $EWSDLLPath $EWSDLL))) {
                Write-Error 'This script requires EWS Managed API 1.2 installed or the DLL in the current folder.'
                Write-Error 'You can download and install EWS Managed API from http://go.microsoft.com/fwlink/?LinkId=255472'
                Exit $ERR_EWSDLLNOTFOUND
            }
        }
        Write-Verbose ('Loading {0}' -f (Join-Path $EWSDLLPath $EWSDLL))
        try {
            # EX2010
            If(!( Get-Module Microsoft.Exchange.WebServices)) {
                Import-Module (Join-Path $EWSDLLPATH $EWSDLL)
            }
        }
        catch {
            #<= EX2010
            [void][Reflection.Assembly]::LoadFile( (Join-Path $EWSDLLPath $EWSDLL))
        }
        try {
            $Temp= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1
        }
        catch {
            Write-Error ('Problem loading {0}' -f $EWSDLL)
            Exit $ERR_EWSLOADING
        }
        Write-Verbose ('Loaded Microsoft.Exchange.WebServices v{0}' -f (Get-Module Microsoft.Exchange.WebServices).Version)
    }
        
    # After calling this any SSL Warning issues caused by Self Signed Certificates will be ignored
    # Source: http://poshcode.org/624
    Function set-TrustAllWeb() {
        Write-Verbose "Set to trust all certificates"
        $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider  
        $Compiler=$Provider.CreateCompiler()  
        $Params=New-Object System.CodeDom.Compiler.CompilerParameters  
        $Params.GenerateExecutable=$False  
        $Params.GenerateInMemory=$True  
        $Params.IncludeDebugInformation=$False  
        $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null  
  
        $TASource= @'
            namespace Local.ToolkitExtensions.Net.CertificatePolicy { 
                public class TrustAll : System.Net.ICertificatePolicy { 
                    public TrustAll() {  
                    }
                    public bool CheckValidationResult(System.Net.ServicePoint sp, System.Security.Cryptography.X509Certificates.X509Certificate cert,   System.Net.WebRequest req, int problem) { 
                        return true; 
                    } 
                } 
            }
'@

        $TAResults=$Provider.CompileAssemblyFromSource($Params, $TASource)  
        $TAAssembly=$TAResults.CompiledAssembly  
        $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")  
        [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll  
    }

    Function Clear-AutoCompleteStream( $EwsService, $EmailAddress, $Pattern) {
        Write-Host ('Processing AutoComplete stream for {0}' -f $EmailAddress)
        $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::Inbox, $EmailAddress)  
        $InboxFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $FolderId)
        $ItemSearchFilterCollection= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.Configuration.AutoComplete")
        $ItemView= New-Object Microsoft.Exchange.WebServices.Data.ItemView( 1)
        $ItemView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet( [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $ItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        $ItemSearchResults= $InboxFolder.FindItems( $ItemSearchFilterCollection, $ItemView)
        If( $ItemSearchResults.Items.Count -gt 0) {
            ForEach( $Item in $ItemSearchResults.Items) {
                If ($pscmdlet.ShouldProcess( 'AutoComplete Stream', 'Clear')) {
                    $res= $Item.Delete( [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
                    Write-Host 'Cleared AutoComplete stream' -Foreground Green
                }
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
            $UserConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($EwsService, "OWA.AutocompleteCache", $FolderId, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
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
        
    ##################################################
    # Main
    ##################################################

    #Requires -Version 2.0

    Load-EWSManagedAPIDLL
    set-TrustAllWeb

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

    ForEach( $ThisIdentity in $Identity) {

        $EmailAddress= get-EmailAddress $ThisIdentity

        If( !$EmailAddress) {
            Write-Error ('Specified mailbox {0} not found' -f $ThisIdentity)
            Exit $ERR_MAILBOXNOTFOUND
        }
        Write-Host ('Processing mailbox {0}' -f $EmailAddress)

        $ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
        $EwsService= New-Object Microsoft.Exchange.WebServices.Data.ExchangeService( $ExchangeVersion)
        $EwsService.UseDefaultCredentials= $true

        If( $Credentials) {
            try {
                Write-Verbose ('Using credentials {0}' -f $Credentials.UserName)
                $EwsService.Credentials= New-Object System.Net.NetworkCredential( $Credentials.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Credentials.Password )))
            }
            catch {
                Write-Error ('Invalid credentials provided {0}' -f $error[0])
                Exit $ERR_INVALIDCREDENTIALS
            }
        }
        Else {
            $EwsService.UseDefaultCredentials= $true
        }

        If( $Impersonation) {
            Write-Verbose ('Using {0} for impersonation' -f $EmailAddress)
            $EwsService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
            $EwsService.HttpHeaders.Add("X-AnchorMailbox", $EmailAddress)
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
                Write-Error ('Autodiscover failed: {0}' -f $error[0])
                Exit $ERR_AUTODISCOVERFAILED
            }
            $ErrorActionPreference= 'Continue'
            Write-Verbose 'Using EWS on CAS {0}' -f $EwsService.Url
        } 
        
        try {
            $RootFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::MsgFolderRoot)
        }
        catch {
            Write-Error ('Can''t access mailbox information store')
            Exit $ERR_CANTACCESSMAILBOXSTORE
        }

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
<#
    .SYNOPSIS
    Clear-AutoComplete
   
    Michel de Rooij
    michel@eightwone.com
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.21, October 27th, 2019
    
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

    To do list:
    - Remove selected entries using pattern matching

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

[cmdletbinding(
    SupportsShouldProcess=$true,
    ConfirmImpact="High"
    )]
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
    [array]$Type= @("Outlook","OWA")
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
        $EWSDLL= "Microsoft.Exchange.WebServices.dll"
        If( Test-Path "$pwd\$EWSDLL") {
            $EWSDLLPath= "$pwd"
        }
        Else {
            $EWSDLLPath = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory'))
            if (!( Test-Path "$EWSDLLPath\$EWSDLL")) {
                Write-Error "This script requires EWS Managed API 1.2 installed or the DLL in the current folder."
                Write-Error "You can download and install EWS Managed API from http://go.microsoft.com/fwlink/?LinkId=255472"
                Exit $ERR_EWSDLLNOTFOUND
            }
        }
        Write-Verbose "Loading $EWSDLLPath\$EWSDLL"
        try {
            # EX2010
            If(!( Get-Module Microsoft.Exchange.WebServices)) {
                Import-Module "$EWSDLLPATH\$EWSDLL"
            }
        }
        catch {
            #<= EX2010
            [void][Reflection.Assembly]::LoadFile( "$EWSDLLPath\$EWSDLL")
        }
        try {
            $Temp= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1
        }
        catch {
            Write-Error "Problem loading $EWSDLL"
            Exit $ERR_EWSLOADING
        }
        Write-Verbose "Loaded Microsoft.Exchange.WebServices v$((Get-Module Microsoft.Exchange.WebServices).Version)"
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

    Function Clear-AutoCompleteStream( $EwsService, $EmailAddress) {
        $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::Inbox, $EmailAddress)  
        $InboxFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $FolderId)
        $ItemSearchFilterCollection= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "IPM.Configuration.AutoComplete")
        $ItemView= New-Object Microsoft.Exchange.WebServices.Data.ItemView( 1)
        $ItemView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet( [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $ItemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        $ItemSearchResults= $InboxFolder.FindItems( $ItemSearchFilterCollection, $ItemView)
        If( $ItemSearchResults.Items.Count -gt 0) {
            ForEach( $Item in $ItemSearchResults.Items) {
                try {
                    If ($pscmdlet.ShouldProcess( 'AutoComplete Stream', 'Clear')) {
                        $res= $Item.Delete( [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
                    }
                    Write-Verbose "Cleared AutoComplete Stream"
                }
                catch {
                    Write-Warning "Problem removing Autocomplete Stream item: " $error[0]
                }
            }
        }
        Else {
            Write-Verbose "No AutoComplete Stream item found."
        }
    }

    Function Clear-OWAAutoComplete( $EwsService, $EmailAddress) {
        Try { 
            $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $EmailAddress)  
            $UserConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($EwsService, "OWA.AutocompleteCache", $FolderId, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
        }
        Catch {
            Write-Verbose "No OWA AutoComplete item found."
        }
        If( $UserConfig) {
            Try {
                If ($pscmdlet.ShouldProcess( 'OWA AutoComplete', 'Clear')) {
                    $UserConfig.Delete()
                    $UserConfig.Update()      
                }
                Write-Verbose 'Cleared OWA AutoComplete'
            }
            Catch {
                Write-Warning "Problem removing OWA Autocomplete Stream item: " $error[0]
            }
        }
    }

    Function Clear-SuggestedContacts( $EwsService, $EmailAddress) {
        $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::MsgFolderRoot, $EmailAddress)  
        $MsgFolderRootFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $FolderId)
        $FolderView= New-Object Microsoft.Exchange.WebServices.Data.FolderView( 1)
        $FolderView.Traversal= [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
        $FolderSearchFilter= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo( [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, "Suggested Contacts")
        $FolderSearchResults= $EwsService.FindFolders( $MsgFolderRootFolder.Id, $FolderSearchFilter, $FolderView)
        If( $FolderSearchResults.Count -gt 0) {
            ForEach( $Folder in $FolderSearchResults) {
                Try {
                    If ($pscmdlet.ShouldProcess( 'SuggestedContacts', 'Clear')) {
                        $Folder.Empty( [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
                    }
                    Write-Verbose "Cleared SuggestedContacts"
                }
                Catch {
                    Write-Error "Problem removing 'Suggested Contacts' folder: " $error[0]
                }
            }
        }
        Else {
            Write-Verbose "No Suggested Contacts folder found."
        }        
    }

    Function Clear-RecipientCache( $EwsService, $EmailAddress) {
        Try {
            $FolderId= New-Object Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecipientCache, $EmailAddress)  
            $RecipientCacheFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $FolderId)
        }
        Catch {
            Write-Verbose "No RecipientCache folder found."
        }
        If( $RecipientCacheFolder) {
            If ($pscmdlet.ShouldProcess( 'RecipientCache', 'Clear')) {
                $RecipientCacheFolder.Empty( [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete)
            }
            Write-Verbose "Cleared RecipientCache"
        }
    }
        
    ##################################################
    # Main
    ##################################################

    #Requires -Version 2.0

    Load-EWSManagedAPIDLL

    If( $Identity -is [array]) {
        # When multiple mailboxes are specified, call script for each mailbox
        [Void]$PSBoundParameters.Remove("Identity")
        $Identity | ForEach-Object { Remove-MessageClassItems -Identity $_ @PSBoundParameters }
    }
    else {
        $EmailAddress= get-EmailAddress $Identity
        If( !$EmailAddress) {
            Write-Error "Specified mailbox $Identity not found"
            Exit $ERR_MAILBOXNOTFOUND
        }
        Write-Host "Processing mailbox $Identity ($EmailAddress)"

        set-TrustAllWeb

        $ExchangeVersion= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1
        $EwsService= New-Object Microsoft.Exchange.WebServices.Data.ExchangeService( $ExchangeVersion)
        $EwsService.UseDefaultCredentials= $true

        If( $Credentials) {
            try {
                Write-Verbose "Using credentials $($Credentials.UserName)"
                $EwsService.Credentials= New-Object System.Net.NetworkCredential( $Credentials.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Credentials.Password )))
            }
            catch {
                Write-Error "Invalid credentials provided " $error[0]
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
            $EwsUrl= "https://$Server/EWS/Exchange.asmx"
            Write-Verbose "Using Exchange Web Services URL $EwsUrl"
            $EwsService.Url= "$EwsUrl"
        }
        Else {
            Write-Verbose "Looking up EWS URL using Autodiscover for $EmailAddress"
            try {
                # Set script to terminate on all errors (autodiscover failure isn't) to make try/catch work
                $ErrorActionPreference= "Stop"
                $EwsService.autodiscoverUrl( $EmailAddress, {$true})
            }
            catch {
                Write-Error "Autodiscover failed: " $error[0]
                Exit $ERR_AUTODISCOVERFAILED
            }
            $ErrorActionPreference= "Continue"
            Write-Verbose "Using EWS on CAS $($EwsService.Url)"
        } 
        
        try {
            $RootFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::MsgFolderRoot)
        }
        catch {
            Write-Error "Can't access mailbox information store"
            Exit $ERR_CANTACCESSMAILBOXSTORE
        }

        If( $Type -contains "All" -or $Type -contains "Outlook") {
            Clear-AutoCompleteStream $EwsService $EmailAddress
        }

        If( $Type -contains "All" -or $Type -contains "OWA") {
            Clear-OWAAutoComplete $EwsService $EmailAddress
        }

        If( $Type -contains "All" -or $Type -contains "SuggestedContacts") {
            Clear-SuggestedContacts $EwsService $EmailAddress
        }

        If( $Type -contains "All" -or $Type -contains "RecipientCache") {
            Clear-RecipientCache $EwsService $EmailAddress
        }
    }   
}   
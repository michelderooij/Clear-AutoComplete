# Clear-AutoComplete

This script allows you to clear one or more locations where recipient information 
is cached, as this could influence end user experience with certain migration scenarios.
You have the option to clear the AutoComplete stream (Name Cache) for Outlook and 
OWA, the Suggested Contacts or the Recipient Cache (Exchange 2013 only).

## Prerequisites

Script requires Microsoft Exchange Web Services (EWS) Managed API 1.2 or up, and Exchange 2010 or up or Exchange Online.
	
## Usage

```
Clear-AutoComplete.ps1 -Mailbox User1 -Type All -Verbose
```
Removes all autocomplete information for mailbox User1.

```
$Credentials= Get-Credential
Clear-AutoComplete.ps1 -Identity olrik@office365tenant.com -Credentials $Credentials 
```

Get credentials and removes Auto Complete information from olrik@office365tenant.com's mailbox.

````
Import-CSV users.csv1 | Clear-AutoComplete.ps1 -Impersonation
````
Uses a CSV file to removes AutoComplete information for a set of mailboxes, using impersonation.

## Contributing

N/A

## Versioning

Initial version published on GitHub is 1.0. Changelog is contained in the script.

## Authors

* Michel de Rooij [initial work] https://github.com/michelderooij

## License

This project is licensed under the MIT License - see the LICENSE.md for details.

## Acknowledgments

N/A
 

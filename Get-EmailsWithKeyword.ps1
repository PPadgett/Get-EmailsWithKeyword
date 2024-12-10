<#
.SYNOPSIS
   Searches Office 365 emails for specific keywords, optionally within a date range, using Microsoft Graph API.

.DESCRIPTION
   This function authenticates with Microsoft Graph API interactively (or uses an existing session) 
   and retrieves emails containing specified keywords in their subject. The search can optionally 
   be filtered by a start and end date. It outputs custom PowerShell objects with sender, subject, and date sent.

.EXAMPLE
   Get-EmailsWithKeyword -Keywords "application", "update"
   Retrieves all emails with "application" or "update" in the subject.

.EXAMPLE
   Get-EmailsWithKeyword -Keywords "invoice", "billing" -StartDate "2024-11-01" -EndDate "2024-11-30"
   Searches emails for "invoice" or "billing" received in November 2024.

.INPUTS
   [string[]] Keywords
   [datetime] StartDate (Optional)
   [datetime] EndDate (Optional)

.OUTPUTS
   [PSCustomObject]
   Custom objects containing Sender, SubjectTitle, and DateSent fields.

.NOTES
   Author: Preston Padgett
   Linkedin: https://www.linkedin.com/in/preston-padgett/
   Date: 2024-12-09

.COMPONENT
   Email Query System

.ROLE
   Email Search Tool

.FUNCTIONALITY
   Retrieves filtered email data from Office 365 using Microsoft Graph API.
#>
[CmdletBinding()]
[OutputType([PSCustomObject])]
Param (
    # Keywords to search in the subject line of emails
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Keywords,

    # Optional start date for filtering emails
    [Parameter(Mandatory = $false, Position = 1)]
    [datetime]$StartDate,

    # Optional end date for filtering emails
    [Parameter(Mandatory = $false, Position = 2)]
    [datetime]$EndDate
)
function Get-AllMailFolders {
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseUri
    )

    $AllFolders = @()
    $FolderUri = "$BaseUri/me/mailFolders"
    do {
        $FolderResponse = Invoke-MgGraphRequest -Method GET -Uri $FolderUri -ErrorAction Stop
        if ($FolderResponse.value) {
            $AllFolders += $FolderResponse.value
        }
        $FolderUri = $FolderResponse.'@odata.nextLink'
    } while ($FolderUri)

    return $AllFolders
}

function Get-FolderMessages {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderId,
        [Parameter(Mandatory=$true)]
        [string]$BaseUri,
        [Parameter(Mandatory=$true)]
        [string]$FullFilter
    )

    $AllFolderMessages = @()
    $MessageUri = "$BaseUri/me/mailFolders/$FolderId/messages?`$filter=$FullFilter"

    do {
        $MessageResponse = Invoke-MgGraphRequest -Method GET -Uri $MessageUri -ErrorAction Stop
        if ($MessageResponse.value) {
            $AllFolderMessages += $MessageResponse.value
        }
        $MessageUri = $MessageResponse.'@odata.nextLink'
    } while ($MessageUri)

    return $AllFolderMessages
}

function Get-EmailsWithKeyword {
    [CmdletBinding(DefaultParameterSetName = 'Default', SupportsShouldProcess = $true, PositionalBinding = $false, ConfirmImpact = 'Low')]
    [Alias("Search-Emails")]
    [OutputType([PSCustomObject])]
    Param (
        # Keywords or key phrases to search in the subject line of emails
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Keywords,

        # Optional start date for filtering emails
        [Parameter(Mandatory = $false, Position = 1)]
        [datetime]$StartDate,

        # Optional end date for filtering emails
        [Parameter(Mandatory = $false, Position = 2)]
        [datetime]$EndDate
    )

    Begin {
        $moduleName = "Microsoft.Graph.Mail"
        if (Get-Module -Name $moduleName -ListAvailable) {
            Write-Host "Module '$moduleName' is installed (available)."
        } else {
            Write-Error "The $moduleName module is not installed. Install it using 'Install-Module -Name Microsoft.Graph.Mail' and try again."
            return
        }

        # Import the module if not already imported
        if (-not (Get-Module -Name Microsoft.Graph.Mail)) {
            Write-Verbose "The Microsoft.Graph.Mail module is not currently imported. Importing it now."
            Import-Module Microsoft.Graph.Mail -ErrorAction Stop
        } else {
            Write-Verbose "The Microsoft.Graph.Mail module is already imported."
        }
    }

    Process {
        try {
            Write-Verbose "Checking for an existing Microsoft Graph session."
            if (-not (Get-MgContext)) {
                Write-Verbose "No active session detected. Connecting to Microsoft Graph interactively."
                Connect-MgGraph -Scopes "Mail.Read"
            } else {
                Write-Verbose "Reusing an existing Microsoft Graph session."
            }

            Write-Verbose "Constructing the email query filter."
            # Build an OR condition for all keywords
            $KeywordFilter = ($Keywords | ForEach-Object { "contains(subject,'$_')" }) -join " or "

            # Construct the date filter if provided
            if ($PSBoundParameters.ContainsKey('StartDate') -and $PSBoundParameters.ContainsKey('EndDate')) {
                Write-Verbose "Applying date range filter from $StartDate to $EndDate."
                $DateFilter = "receivedDateTime ge $($StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ')) and receivedDateTime le $($EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
                $FullFilter = "($KeywordFilter) and $DateFilter"
            } elseif ($PSBoundParameters.ContainsKey('StartDate')) {
                Write-Verbose "Applying start date filter from $StartDate."
                $DateFilter = "receivedDateTime ge $($StartDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
                $FullFilter = "($KeywordFilter) and $DateFilter"
            } elseif ($PSBoundParameters.ContainsKey('EndDate')) {
                Write-Verbose "Applying end date filter up to $EndDate."
                $DateFilter = "receivedDateTime le $($EndDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
                $FullFilter = "($KeywordFilter) and $DateFilter"
            } else {
                Write-Verbose "No date filters applied. Using only keyword filter."
                $FullFilter = "($KeywordFilter)"
            }

            Write-Verbose "Generated query filter: $FullFilter"

            # Now enumerate all mail folders to ensure we cover all (including Junk)
            $BaseUri = "https://graph.microsoft.com/v1.0"
            Write-Verbose "Retrieving all mail folders."
            $AllFolders = Get-AllMailFolders -BaseUri $BaseUri
            
            # Create a lookup hashtable of FolderId to DisplayName
            $FolderLookup = @{}
            foreach ($Folder in $AllFolders) {
                $FolderLookup[$Folder.id] = $Folder.displayName
            }

            Write-Verbose "Total folders found: $($AllFolders.Count). Enumerating messages in each folder."
            $AllEmails = @()
            foreach ($Folder in $AllFolders) {
                Write-Verbose "Retrieving messages from folder: $($Folder.displayName) (ID: $($Folder.id))"
                $FolderMessages = Get-FolderMessages -FolderId $Folder.id -BaseUri $BaseUri -FullFilter $FullFilter
                if ($FolderMessages) {
                    $AllEmails += $FolderMessages
                }
            }

            # Debug: Check what we got before filtering
            if ($AllEmails.Count -gt 0) {
                Write-Verbose "Raw email subjects returned from all folders:"
                $AllEmails | ForEach-Object { Write-Verbose $_.subject }
            } else {
                Write-Verbose "No emails returned from Graph for the given filters."
            }

            Write-Verbose "Processing the retrieved email data on the client side."

            # Check each message against all keywords in a case-insensitive manner
            $FilteredEmails = $AllEmails | Where-Object {
                $Subject = $_.subject -as [string]
                if (-not $Subject) { return $false }

                $LowerSubject = $Subject.ToLower()

                # Instead of breaking at the first match, we check all
                $Matches = $false
                foreach ($Keyword in $Keywords) {
                    $LowerKeyword = $Keyword.ToLower()
                    if ($LowerSubject -like "*$LowerKeyword*") {
                        $Matches = $true
                    }
                }
                $Matches
            } | ForEach-Object {
                $FolderName = $FolderLookup[$_.parentFolderId]
                [PSCustomObject]@{
                    Sender       = $_.sender.emailAddress.name
                    SubjectTitle = $_.subject
                    DateSent     = $_.sentDateTime
                    Folder       = $FolderName
                }
            }

            $Count = $FilteredEmails.Count
            Write-Host "Total emails found with the specified keywords: $Count"
            return $FilteredEmails
        } catch {
            Write-Error "An error occurred while querying Microsoft Graph API: $_"
        }
    }

    End {
        Write-Verbose "Completed the email search operation."
    }
}



# Check if the script is being executed directly with parameters, if not is assumed to be dot-sourced for testing
if ($MyInvocation.MyCommand.CommandType -eq 'ExternalScript' -and $MyInvocation.MyCommand.Name -and $PSBoundParameters.Count -gt 0) {
    Write-Verbose "Identified parameter set: $($PSCmdlet.ParameterSetName)"

    # Execute the function based on the parameter set used passed to the script.
    switch ($PSCmdlet.ParameterSetName) {
        '__AllParameterSets' {
            Write-Verbose "Executing Invoke-PTUnit function based on parameter set"
            Get-EmailsWithKeyword @PSBoundParameters
        }
    }
}
else {
    Write-Verbose "Skipping function execution due to the absence of required execution context or parameters. Script Function has been dot-sourced"
}
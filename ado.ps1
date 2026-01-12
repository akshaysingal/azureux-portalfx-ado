#==============================================================================
# AZURE DEVOPS WORK ITEM MANAGEMENT SCRIPT
#==============================================================================

#------------------------------------------------------------------------------
# HELPER FUNCTIONS (Internal use only)
#------------------------------------------------------------------------------

function Ensure-AzDevOpsLogin {
    try {
        az account show --only-show-errors | Out-Null
    }
    catch {
        Write-Host "Azure login required. Launching az login..." -ForegroundColor Yellow
        az login | Out-Null
    }

    try {
        az devops project list --only-show-errors | Out-Null
    }
    catch {
        Write-Host "Azure DevOps login required. Launching az devops login..." -ForegroundColor Yellow
        az devops login | Out-Null
    }
}

function New-AdoWorkItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Title,

        [Parameter()]
        [string] $Type = 'Bug',

        [Parameter()]
        [string] $State = 'Active',

        [Parameter()]
        [string] $Organization = 'https://dev.azure.com/msazure',

        [Parameter()]
        [string] $Project = 'One',

        [Parameter()]
        [string] $Area = 'One\Azure Portal\Hubs',

        [Parameter()]
        [string] $Iteration = 'One\Krypton',

        [Parameter()]
        [string] $AssignedTo = 'aksingal@microsoft.com',

        [Parameter()]
        [string] $Tags = 'autogen',

        [Parameter()]
        [string] $ApiVersion = '7.1',

        [Parameter()]
        [switch] $QuietlyReturn
    )

    $patch = @(
        @{ op = 'add'; path = '/fields/System.Title';         value = $Title }
        @{ op = 'add'; path = '/fields/System.AreaPath';      value = $Area }
        @{ op = 'add'; path = '/fields/System.IterationPath'; value = $Iteration }
        @{ op = 'add'; path = '/fields/System.AssignedTo';    value = $AssignedTo }
        @{ op = 'add'; path = '/fields/System.State';         value = $State }
        @{ op = 'add'; path = '/fields/System.Tags';          value = $Tags }
    ) | ConvertTo-Json -Depth 10

    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        $patch | Out-File -FilePath $tempFile -Encoding utf8

        $attempt = 0
        $maxAttempts = 2

        do {
            try {
                $response = az devops invoke `
                    --organization $Organization `
                    --area wit `
                    --resource workitems `
                    --route-parameters "project=$Project" "type=$Type" `
                    --http-method POST `
                    --api-version $ApiVersion `
                    --media-type "application/json-patch+json" `
                    --in-file $tempFile `
                    --only-show-errors |
                    ConvertFrom-Json

                break
            }
            catch {
                if ($attempt -eq 0) {
                    Write-Host "Auth may be expired. Attempting login..." -ForegroundColor Yellow
                    Ensure-AzDevOpsLogin
                }
                else {
                    throw
                }
            }

            $attempt++
        }
        while ($attempt -lt $maxAttempts)
    }
    finally {
        Remove-Item $tempFile -ErrorAction SilentlyContinue
    }

    $retVal = "#$($response.id): $($response.fields.'System.Title')"

    if ($QuietlyReturn) {
        return $retVal
    }

    Write-Host "$Type $retVal ($State)" -ForegroundColor Green
    Write-Host "https://msazure.visualstudio.com/One/_workitems/edit/$($response.id)" -ForegroundColor Green
}

function New-WorkItemInternal {
    <#
    .SYNOPSIS
        Generic work item creation helper.
    .DESCRIPTION
        Internal helper function that maps work item types and states for the public wrapper functions.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Active')]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateSet('bug','pbi')]
        [string] $Kind,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string] $Title,

        [Parameter(ParameterSetName='New')]
        [switch] $New,

        [Parameter(ParameterSetName='Active')]
        [switch] $Active,

        [Parameter(ParameterSetName='InReview')]
        [switch] $InReview,

        [Parameter(ParameterSetName='Committed')]
        [switch] $Committed,

        [Parameter(ParameterSetName='Approved')]
        [switch] $Approved
    )

    $typeMap = @{
        bug = 'Bug'
        pbi = 'Product Backlog Item'
    }

    $stateKey =
        if     ($New)       { 'new' }
        elseif ($InReview)  { 'inreview' }
        elseif ($Committed) { 'committed' }
        elseif ($Approved)  { 'approved' }
        else                { 'active' }

    $stateMap = @{
        active    = 'Active'
        approved  = 'Approved'
        committed = 'Committed'
        inreview  = 'In Review'
        new       = 'New'
    }

    New-AdoWorkItem -Type $typeMap[$Kind] -Title $Title.Trim() -State $stateMap[$stateKey] -QuietlyReturn:$false
}

#------------------------------------------------------------------------------
# PUBLIC WRAPPER FUNCTIONS (User-facing commands)
#------------------------------------------------------------------------------

function CreateWorkItem {
    <#
    .SYNOPSIS
        Interactive work item creation with user prompts.
    .DESCRIPTION
        Prompts the user for all work item details and creates a new work item in Azure DevOps.
        Uses default values when user provides empty input (except for Title which is mandatory).
    .EXAMPLE
        CreateWorkItem
    #>
    [CmdletBinding()]
    param()

    Write-Host "`n=== Azure DevOps Work Item Creation ===" -ForegroundColor Cyan
    Write-Host "Press Enter to use default values (shown in brackets)" -ForegroundColor Gray
    Write-Host ""

    # Prompt for Kind with validation
    do {
        $kind = Read-Host "Work Item Type [Bug]: (Bug/PBI/Task)"
        if ([string]::IsNullOrWhiteSpace($kind)) {
            $kind = "Bug"
        }
        $kind = $kind.Trim()
        
        if ($kind -notin @('Bug', 'PBI', 'Product Backlog Item', 'Task')) {
            Write-Host "Invalid work item type. Please enter 'Bug', 'PBI', or 'Task'." -ForegroundColor Red
            $kind = $null
        }
        elseif ($kind -eq 'PBI') {
            $kind = 'Product Backlog Item'
        }
    } while ([string]::IsNullOrWhiteSpace($kind))

    # Prompt for Title - mandatory field
    do {
        $title = Read-Host "Title (required)"
        if ([string]::IsNullOrWhiteSpace($title)) {
            Write-Host "Title cannot be empty. Please provide a title." -ForegroundColor Red
        }
    } while ([string]::IsNullOrWhiteSpace($title))

    # Prompt for Area Path
    $areaPath = Read-Host "Area Path [One\Azure Portal\Hubs]"
    if ([string]::IsNullOrWhiteSpace($areaPath)) {
        $areaPath = "One\Azure Portal\Hubs"
    }

    # Prompt for Iteration Path
    $iterationPath = Read-Host "Iteration Path [One\Krypton]"
    if ([string]::IsNullOrWhiteSpace($iterationPath)) {
        $iterationPath = "One\Krypton"
    }

    # Prompt for Assigned To
    $assignedTo = Read-Host "Assigned To [aksingal]"
    if ([string]::IsNullOrWhiteSpace($assignedTo)) {
        $assignedTo = "aksingal"
    }
    # Add @microsoft.com suffix
    $assignedTo = "$assignedTo@microsoft.com"

    # Prompt for State with validation
    do {
        $state = Read-Host "State [Active]: (New/Active/Committed/Approved/In Review)"
        if ([string]::IsNullOrWhiteSpace($state)) {
            $state = "Active"
        }
        $state = $state.Trim()
        
        # Convert to proper case for validation and storage
        $validStates = @{
            'new' = 'New'
            'active' = 'Active'
            'committed' = 'Committed'
            'approved' = 'Approved'
            'in review' = 'In Review'
        }
        
        $stateLower = $state.ToLower()
        if ($validStates.ContainsKey($stateLower)) {
            $state = $validStates[$stateLower]
        } else {
            Write-Host "Invalid state. Please enter: New, Active, Committed, Approved, or 'In Review'." -ForegroundColor Red
            $state = $null
        }
    } while ([string]::IsNullOrWhiteSpace($state))

    # Prompt for Tags
    $additionalTags = Read-Host "Additional Tags (will be appended to 'autogen')"
    if ([string]::IsNullOrWhiteSpace($additionalTags)) {
        $tags = "autogen"
    } else {
        $tags = "autogen; $($additionalTags.Trim())"
    }

    # Display summary and confirm
    Write-Host "`n=== Work Item Summary ===" -ForegroundColor Yellow
    Write-Host "Type: $kind"
    Write-Host "Title: $title"
    Write-Host "Area Path: $areaPath"
    Write-Host "Iteration Path: $iterationPath"
    Write-Host "Assigned To: $assignedTo"
    Write-Host "State: $state"
    Write-Host "Tags: $tags"
    Write-Host ""

    $confirm = Read-Host "Create this work item? (y/N)"
    if ($confirm -match '^[Yy]') {
        try {
            # Create the work item
            $result = New-AdoWorkItem -Type $kind -Title $title.Trim() -Area $areaPath -Iteration $iterationPath -AssignedTo $assignedTo -State $state -Tags $tags -QuietlyReturn:$false
            Write-Host "`nWork item created successfully!" -ForegroundColor Green
        }
        catch {
            Write-Host "`nFailed to create work item: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Work item creation cancelled." -ForegroundColor Yellow
    }
}

function bug {
    <#
    .SYNOPSIS
        Creates a bug work item.
    .DESCRIPTION
        Shorthand command to create bug work items with various states.
    .PARAMETER Title
        Title of the bug
    .EXAMPLE
        bug "Login fails on IE" -Active
        bug "Memory leak in service" -New
    #>
    [CmdletBinding(DefaultParameterSetName = 'Active')]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string] $Title,

        [Parameter(ParameterSetName='New')]
        [switch] $New,

        [Parameter(ParameterSetName='Active')]
        [switch] $Active,

        [Parameter(ParameterSetName='InReview')]
        [switch] $InReview,

        [Parameter(ParameterSetName='Committed')]
        [switch] $Committed,

        [Parameter(ParameterSetName='Approved')]
        [switch] $Approved
    )

    New-WorkItemInternal -Kind bug -Title $Title @PSBoundParameters
}

function pbi {
    <#
    .SYNOPSIS
        Creates a Product Backlog Item (PBI) work item.
    .DESCRIPTION
        Shorthand command to create PBI work items with various states.
    .PARAMETER Title
        Title of the PBI
    .EXAMPLE
        pbi "Add user management feature" -New
        pbi "Implement search functionality" -Active
    #>
    [CmdletBinding(DefaultParameterSetName = 'Active')]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string] $Title,

        [Parameter(ParameterSetName='New')]
        [switch] $New,

        [Parameter(ParameterSetName='Active')]
        [switch] $Active,

        [Parameter(ParameterSetName='InReview')]
        [switch] $InReview,

        [Parameter(ParameterSetName='Committed')]
        [switch] $Committed,

        [Parameter(ParameterSetName='Approved')]
        [switch] $Approved
    )

    New-WorkItemInternal -Kind pbi -Title $Title @PSBoundParameters
}

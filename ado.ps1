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

    $confirm = Read-Host "Create this work item? (y/n)"
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

function pr {
    <#
    .SYNOPSIS
        Creates a bug work item, commits staged files, and creates a pull request.
    .DESCRIPTION
        Comprehensive workflow function that:
        1. Creates a new bug work item with user-provided title
        2. Commits all staged files using the work item ID and title as commit message
        3. Creates a pull request for the changes
    .PARAMETER Title
        Title for the bug work item (will also be used in commit message)
    .EXAMPLE
        CreatePR
    #>
    [CmdletBinding()]
    param()

    Write-Host "`n=== Bug + PR Creation Workflow ===" -ForegroundColor Cyan
    Write-Host ""

    # First prompt: Ask if user wants to create a work item
    do {
        $createWorkItem = Read-Host "Do you want to create a new work item? (y/n)"
        $createWorkItem = $createWorkItem.Trim().ToLower()
        
        if ([string]::IsNullOrWhiteSpace($createWorkItem) -or $createWorkItem -eq 'n') {
            $createWorkItem = $false
            break
        }
        elseif ($createWorkItem -eq 'y') {
            $createWorkItem = $true
            break
        }
        else {
            Write-Host "Please enter 'y' for yes or 'n' for no." -ForegroundColor Red
        }
    } while ($true)

    if ($createWorkItem) {
        # Work item creation flow
        Write-Host "`n--- Work Item Creation ---" -ForegroundColor Cyan
        
        # Prompt for work item kind
        do {
            $workItemKind = Read-Host "Work Item Type [Bug]: (Bug/PBI)"
            if ([string]::IsNullOrWhiteSpace($workItemKind)) {
                $workItemKind = "Bug"
            }
            $workItemKind = $workItemKind.Trim()
            
            if ($workItemKind -notin @('Bug', 'PBI', 'Product Backlog Item')) {
                Write-Host "Invalid work item type. Please enter 'Bug' or 'PBI'." -ForegroundColor Red
                $workItemKind = $null
            }
            elseif ($workItemKind -eq 'PBI') {
                $workItemKind = 'Product Backlog Item'
            }
        } while ([string]::IsNullOrWhiteSpace($workItemKind))

        # Prompt for work item title
        do {
            $title = Read-Host "Work item title (required)"
            if ([string]::IsNullOrWhiteSpace($title)) {
                Write-Host "Title cannot be empty. Please provide a title." -ForegroundColor Red
            }
        } while ([string]::IsNullOrWhiteSpace($title))

        Write-Host "Work Item Type: $workItemKind" -ForegroundColor White
        Write-Host "Title: $title" -ForegroundColor White
    }
    else {
        # PR-only flow
        Write-Host "`n--- Pull Request Creation ---" -ForegroundColor Cyan
        
        # Prompt for PR title
        do {
            $title = Read-Host "Pull request title (required)"
            if ([string]::IsNullOrWhiteSpace($title)) {
                Write-Host "Title cannot be empty. Please provide a title." -ForegroundColor Red
            }
        } while ([string]::IsNullOrWhiteSpace($title))

        Write-Host "PR Title: $title" -ForegroundColor White
    }

    # Prompt for additional reviewers
    Write-Host "`n--- Reviewers ---" -ForegroundColor Cyan
    $defaultReviewers = @('nickkirc', 'bifunk')
    Write-Host "Default reviewers: $($defaultReviewers -join ', ')" -ForegroundColor Gray
    
    $additionalReviewers = Read-Host "Additional reviewers (comma-separated, optional)"
    $reviewers = $defaultReviewers
    
    if (-not [string]::IsNullOrWhiteSpace($additionalReviewers)) {
        $additionalList = $additionalReviewers.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $reviewers += $additionalList
    }
    
    Write-Host "Final reviewers: $($reviewers -join ', ')" -ForegroundColor White

    Write-Host ""

    try {
        $workItemNumber = $null
        $commitMessage = $null

        if ($createWorkItem) {
            # Step 1: Create the work item
            Write-Host "Step 1: Creating work item..." -ForegroundColor Yellow
            $workItemResult = New-AdoWorkItem -Type $workItemKind -Title $title.Trim() -State 'In Review' -QuietlyReturn:$true
            
            # Extract work item number from result (format: "#123: Title")
            if ($workItemResult -match '^#(\d+):') {
                $workItemNumber = $matches[1]
                Write-Host "✓ Work item created: $workItemResult" -ForegroundColor Green
                $commitMessage = "#${workItemNumber}: ${title}"
            } else {
                throw "Failed to parse work item number from result: $workItemResult"
            }
        }
        else {
            # Step 1: Skip work item creation
            Write-Host "Step 1: Skipping work item creation..." -ForegroundColor Yellow
            $commitMessage = $title
            Write-Host "✓ Using PR title for commit message" -ForegroundColor Green
        }

        # Step 2: Check for staged files
        Write-Host "`nStep 2: Checking for staged files..." -ForegroundColor Yellow
        $stagedFiles = git diff --cached --name-only
        if (-not $stagedFiles) {
            Write-Host "⚠ No staged files found. Will create empty commit." -ForegroundColor Yellow
        } else {
            Write-Host "✓ Found $($stagedFiles.Count) staged file(s):" -ForegroundColor Green
            $stagedFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }
        }

        # Step 3: Create commit message and commit
        Write-Host "`nStep 3: Committing staged files..." -ForegroundColor Yellow
        
        # Use --allow-empty flag if no staged files
        if (-not $stagedFiles) {
            $commitResult = git commit --allow-empty -m $commitMessage
        } else {
            $commitResult = git commit -m $commitMessage
        }
        
        if ($LASTEXITCODE -ne 0) {
            throw "Git commit failed with exit code: $LASTEXITCODE"
        }
        Write-Host "✓ Committed with message: $commitMessage" -ForegroundColor Green

        # Step 4: Push changes
        Write-Host "`nStep 4: Pushing changes to remote..." -ForegroundColor Yellow
        $currentBranch = git branch --show-current
        
        $pushResult = git push -f origin $currentBranch
        if ($LASTEXITCODE -ne 0) {
            throw "Git push failed with exit code: $LASTEXITCODE"
        }
        Write-Host "✓ Pushed branch '$currentBranch' to remote" -ForegroundColor Green

        # Step 5: Create pull request
        Write-Host "`nStep 5: Creating pull request..." -ForegroundColor Yellow
        
        # Get repository name from git remote
        $remoteUrl = git remote get-url origin
        if ($remoteUrl -match '/([^/]+)\.git$' -or $remoteUrl -match '/([^/]+)$') {
            $repositoryName = $matches[1]
        } else {
            throw "Could not determine repository name from remote URL: $remoteUrl"
        }
        
        # Get target branch (assume main/master)
        $targetBranch = "dev"

        # Build the PR creation command
        $prArgs = @(
            'repos', 'pr', 'create',
            '--organization', 'https://dev.azure.com/msazure',
            '--project', 'One',
            '--repository', $repositoryName,
            '--source-branch', $currentBranch,
            '--target-branch', $targetBranch,
            '--title', $commitMessage,
            '--description', "Automated PR$(if ($workItemNumber) { " for work item #$workItemNumber" })`n`nChanges: $title",
            '--only-show-errors'
        )

        # Add work item if available
        if ($workItemNumber) {
            $prArgs += @('--work-items', $workItemNumber)
        }

        # Add reviewers
        if ($reviewers -and $reviewers.Count -gt 0) {
            $reviewersWithSuffix = $reviewers | ForEach-Object { "$_@microsoft.com" }
            $prArgs += @('--optional-reviewers', ($reviewersWithSuffix -join ' '))
        }

        $prResult = az @prArgs | ConvertFrom-Json

        if ($prResult) {
            Write-Host "✓ Pull request #$($prResult.pullRequestId) created successfully!" -ForegroundColor Green
        } else {
            throw "Failed to create pull request"
        }

        Write-Host "`n=== Workflow Complete! ===" -ForegroundColor Green
        if ($workItemNumber) {
            Write-Host "Work Item: $workItemResult" -ForegroundColor White
        }
        Write-Host "Commit: $commitMessage" -ForegroundColor White
        $prUrl = "https://msazure.visualstudio.com/One/_git/$repositoryName/pullrequest/$($prResult.pullRequestId)"
        Write-Host "PR: $prUrl" -ForegroundColor White
        
        # Open PR in Edge browser
        Start-Process "msedge" $prUrl

    }
    catch {
        Write-Host "`n❌ Workflow failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "You may need to manually clean up any partially completed steps." -ForegroundColor Yellow
        throw
    }
}

#------------------------------------------------------------------------------
# GIT ALIASES
#------------------------------------------------------------------------------

# Quick git status alias with last commit
function s { 
    git status $args
    Write-Host ""
    Write-Host "Last commit:" -ForegroundColor Yellow
    git log -1 --oneline
}

# New branch creation alias
function nb {
    param([Parameter(Mandatory)]$branchName)
    
    if ([string]::IsNullOrWhiteSpace($branchName)) {
        Write-Error "Branch name cannot be empty or null" -ErrorAction Stop
        return
    }
    
    git checkout -b "aksingal/$branchName" origin/dev

    git pull
    git rebase
}



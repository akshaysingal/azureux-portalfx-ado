#==============================================================================
# AZURE DEVOPS POWERSHELL TOOLKIT
#==============================================================================
# 
# DESCRIPTION:
#   Comprehensive PowerShell toolkit for Azure DevOps operations including:
#   - Work item management (creation, querying, lifecycle management)
#   - Pipeline test execution with UI selection
#   - Extension configuration lookup across multiple cloud environments
#   - Git workflow automation (branch creation, PR management)
#   - Authentication handling with automatic retry logic
#
# AUTHOR: Akshay Singal
#
# DEPENDENCIES:
#   - Azure CLI (az) with DevOps extension
#   - Git command line tools
#   - PowerShell 5.1 or higher
#   - Windows Forms (for UI components)
#
# CONFIGURATION:
#   - Default organization: https://dev.azure.com/msazure
#   - Default project: One
#
# USAGE EXAMPLES:
#   createworkitem          # Interactive work item creation
#   bug "Fix login issue"    # Quick bug creation
#   runtests               # Launch test selection UI
#   createpr               # Create work item + PR workflow
#   listworkitems          # View assigned work items
#   getextensiondetails    # Extension configuration lookup
#
#==============================================================================

#==============================================================================
# CORE HELPER FUNCTIONS
#==============================================================================
# 
# This section contains internal helper functions that provide:
# - Authentication management with automatic retry logic
# - Centralized Azure DevOps API error handling
# - Work item creation and lifecycle management
# - Background task execution with progress indicators
# 
# These functions are designed for internal use by the public wrapper functions
# and should not be called directly by end users.
#==============================================================================

function Ensure-AzDevOpsLogin {
    <#
    .SYNOPSIS
        Ensures the user is authenticated with both Azure CLI and Azure DevOps CLI.
    .DESCRIPTION
        Verifies that the user has valid authentication tokens for both Azure CLI (az) 
        and Azure DevOps CLI (az devops). If either authentication is missing or expired,
        automatically prompts for login. This function is called internally by 
        Invoke-WithAzDevOpsAuth to handle authentication failures gracefully.
    .EXAMPLE
        Ensure-AzDevOpsLogin
        
        Verifies authentication status and prompts for login if needed.
    .NOTES
        - Checks Azure CLI authentication with 'az account show'
        - Checks Azure DevOps authentication with 'az devops project list'
        - Automatically launches login prompts when authentication is required
        - Uses --only-show-errors flag to minimize noise during verification
    #>
    Write-Host "⏱ Verifying authentication..." -ForegroundColor Gray
    
    # Verify Azure CLI authentication status
    # If this fails, the user needs to run 'az login' to authenticate with Azure
    try {
        az account show --only-show-errors | Out-Null
    }
    catch {
        Write-Host "Azure login required. Launching az login..." -ForegroundColor Yellow
        az login | Out-Null
    }

    # Verify Azure DevOps CLI authentication status  
    # If this fails, the user needs to run 'az devops login' for DevOps-specific authentication
    try {
        az devops project list --only-show-errors | Out-Null
    }
    catch {
        Write-Host "Azure DevOps login required. Launching az devops login..." -ForegroundColor Yellow
        az devops login | Out-Null
    }
    
    Write-Host "✓ User already authenticated, continuing..." -ForegroundColor Gray
}

function Invoke-WithAzDevOpsAuth {
    <#
    .SYNOPSIS
        Executes a script block with automatic Azure DevOps authentication retry logic.
    .DESCRIPTION
        Attempts to execute the provided script block. If the operation fails due to 
        authentication errors, it will re-authenticate using Ensure-AzDevOpsLogin and 
        retry the operation once. This eliminates the need to call Ensure-AzDevOpsLogin 
        manually in every function that uses Azure DevOps.
    .PARAMETER ScriptBlock
        The script block to execute that contains Azure DevOps operations.
    .PARAMETER ArgumentList
        Arguments to pass to the script block.
    .EXAMPLE
        $result = Invoke-WithAzDevOpsAuth -ScriptBlock { 
            az devops project list --only-show-errors 
        }
    .EXAMPLE
        $result = Invoke-WithAzDevOpsAuth -ScriptBlock { 
            param($org, $project)
            az devops invoke --organization $org --project $project
        } -ArgumentList $Organization, $ProjectGuid
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ScriptBlock] $ScriptBlock,
        
        [Parameter()]
        [Object[]] $ArgumentList = @()
    )
    
    try {
        # First attempt - execute the script block
        if ($ArgumentList.Count -gt 0) {
            return & $ScriptBlock @ArgumentList
        } else {
            return & $ScriptBlock
        }
    }
    catch {
        $errorMessage = $_.Exception.Message.ToLower()
        
        # Enhanced authentication error detection
        # Check if the error message contains any authentication-related keywords
        $authKeywords = @(
            "authentication", "unauthorized", "login", "token", "credential", 
            "permission denied", "access denied", "401", "403", "az login", 
            "az devops login", "not authenticated", "session has expired", 
            "invalid_token", "token expired", "authentication required"
        )
        
        $isAuthError = $authKeywords | Where-Object { $errorMessage -match $_ } | Measure-Object | ForEach-Object Count
        $isAuthError = $isAuthError -gt 0
        
        if ($isAuthError) {
            Write-Host "Authentication error detected. Re-authenticating..." -ForegroundColor Yellow
            
            # Re-authenticate
            Ensure-AzDevOpsLogin
            
            # Retry the operation once
            try {
                if ($ArgumentList.Count -gt 0) {
                    return & $ScriptBlock @ArgumentList
                } else {
                    return & $ScriptBlock
                }
            }
            catch {
                # If it still fails after re-auth, re-throw the error
                throw
            }
        } else {
            # Not an auth error, re-throw the original error
            throw
        }
    }
}

function New-AdoWorkItem {
    <#
    .SYNOPSIS
        Creates a new work item in Azure DevOps with comprehensive field mapping.
    .DESCRIPTION
        Creates work items (Bug, PBI, Task) in Azure DevOps using the REST API.
        Supports all common work item fields including title, area path, iteration,
        assignment, state, and tags. Uses JSON Patch format for API operations
        and includes automatic authentication handling.
    .PARAMETER Title
        The title/summary of the work item. This is a required field.
    .PARAMETER Type
        Work item type (Bug, Product Backlog Item, Task). Defaults to 'Bug'.
    .PARAMETER State
        Initial state of the work item (New, Active, In Review, etc.). Defaults to 'Active'.
    .PARAMETER Organization
        Azure DevOps organization URL. Defaults to 'https://dev.azure.com/msazure'.
    .PARAMETER Project
        Project name within the organization. Defaults to 'One'.
    .PARAMETER Area
        Area path for categorizing the work item. Defaults to 'One\Azure Portal\Hubs'.
    .PARAMETER Iteration
        Iteration path for sprint/milestone assignment. Defaults to 'One\Krypton'.
    .PARAMETER AssignedTo
        Email address of the person to assign the work item to.
    .PARAMETER Tags
        Semicolon-separated tags for categorization. Defaults to 'autogen'.
    .PARAMETER ApiVersion
        Azure DevOps REST API version. Defaults to '7.1'.
    .PARAMETER QuietlyReturn
        When specified, suppresses console output and only returns the work item reference.
    .OUTPUTS
        String in format "#12345: Work Item Title"
    .EXAMPLE
        New-AdoWorkItem -Title "Fix login bug" -Type "Bug" -State "Active"
        
        Creates a new bug work item with specified title in Active state.
    .EXAMPLE
        $workItem = New-AdoWorkItem -Title "New feature request" -Type "Product Backlog Item" -QuietlyReturn
        
        Creates a PBI and returns the reference without console output.
    .NOTES
        - Uses JSON Patch format for Azure DevOps REST API
        - Automatically handles authentication via Invoke-WithAzDevOpsAuth
        - Creates temporary files for API payload (automatically cleaned up)
        - Returns work item URL for easy access in browser
    #>
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

    # Create JSON Patch payload for work item creation
    # Azure DevOps uses RFC 6902 JSON Patch format for work item operations
    $patch = @(
        @{ op = 'add'; path = '/fields/System.Title';         value = $Title }
        @{ op = 'add'; path = '/fields/System.AreaPath';      value = $Area }
        @{ op = 'add'; path = '/fields/System.IterationPath'; value = $Iteration }
        @{ op = 'add'; path = '/fields/System.AssignedTo';    value = $AssignedTo }
        @{ op = 'add'; path = '/fields/System.State';         value = $State }
        @{ op = 'add'; path = '/fields/System.Tags';          value = $Tags }
    ) | ConvertTo-Json -Depth 10

    # Create temporary file for API payload
    # Azure DevOps CLI requires input from file for complex JSON payloads
    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        $patch | Out-File -FilePath $tempFile -Encoding utf8

        # Execute work item creation with automatic authentication handling
        $response = Invoke-WithAzDevOpsAuth -ScriptBlock {
            az devops invoke `
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
        }
    }
    finally {
        # Always clean up temporary file, even if operation fails
        Remove-Item $tempFile -ErrorAction SilentlyContinue
    }

    # Format return value with work item ID and title
    $retVal = "#$($response.id): $($response.fields.'System.Title')"

    if ($QuietlyReturn) {
        return $retVal
    }

    # Display success message with work item details and URL
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

function Invoke-WithSpinner {
    <#
    .SYNOPSIS
        Executes a script block in the background while showing a spinning animation.
    .DESCRIPTION
        Runs the provided script block as a background job and displays a spinning animation
        with the specified message until the job completes.
    .PARAMETER ScriptBlock
        The script block to execute in the background.
    .PARAMETER ArgumentList
        Arguments to pass to the script block.
    .PARAMETER Message
        The message to display with the spinner (e.g., "Loading data").
    .EXAMPLE
        $result = Invoke-WithSpinner -ScriptBlock { Get-Process } -Message "Loading processes"
    .EXAMPLE
        $result = Invoke-WithSpinner -ScriptBlock { param($name) Get-Service $name } -ArgumentList "Spooler" -Message "Getting service"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ScriptBlock] $ScriptBlock,
        
        [Parameter()]
        [Object[]] $ArgumentList = @(),
        
        [Parameter(Mandatory)]
        [string] $Message
    )
    
    try {
        # Start spinner animation
        $spinnerChars = @('⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏')
        $spinnerIndex = 0
        
        # Start background job
        $job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
        
        # Show spinner while job is running
        while ($job.State -eq 'Running') {
            $spinner = $spinnerChars[$spinnerIndex % $spinnerChars.Length]
            Write-Host "`r$spinner $Message... " -NoNewline -ForegroundColor Yellow
            $spinnerIndex++
            Start-Sleep -Milliseconds 100
        }
        
        # Get the result and clean up
        $result = Receive-Job -Job $job
        Remove-Job -Job $job
        
        return $result
    }
    catch {
        # Clean up job if something goes wrong
        if ($job) {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
        throw
    }
}

#==============================================================================
# EXTENSION CONFIGURATION MANAGEMENT
#==============================================================================
# 
# This section provides powerful extension configuration lookup capabilities
# across multiple Azure cloud environments. The system maintains in-memory
# caches of extension configurations from various appsettings.{cloud}.json
# files in the AzureUX-ExtensionStudio repository.
#
# KEY FEATURES:
# - Multi-cloud extension configuration lookup (production, fairfax, mooncake, etc.)
# - Dual indexing: by extension name (PortalName) and OAuth Client ID
# - Parallel fetching for optimal performance
# - Interactive search with suggestions and partial matching
# - Automatic cache initialization and refresh capabilities
# - Hyperlink generation for ICM on-call and 1P app configuration
#
# SUPPORTED CLOUDS:
# - production (prod) - Main production environment
# - fairfax (ff) - US Government cloud  
# - mooncake (mc) - China cloud
# - bleu - French cloud
# - usnat/ussec - Classified government clouds
# - delos - Microsoft internal cloud
# - dogfood (df) - Pre-production testing environment
#
#==============================================================================

# Global variables to store the extension caches
# These provide fast O(1) lookup performance for extension configurations
$script:ExtensionCache = $null                    # Indexed by PortalName -> Cloud -> Extension
$script:ExtensionCacheByOAuthClientId = $null     # Indexed by OAuthClientId -> Cloud -> Extension[]
$script:InitializedClouds = @{}                   # Tracks which clouds have been successfully loaded

function Initialize-ExtensionCache {
    <#
    .SYNOPSIS
        Reads all appsettings.{cloud}.json files and caches Service.HostingService.Extensions data.
    .DESCRIPTION
        Fetches all appsettings.{cloud}.json files from the AzureUX-ExtensionStudio repository,
        extracts the Service.HostingService.Extensions arrays, and creates an in-memory lookup
        map indexed by PortalName and cloud for fast retrieval.
    .PARAMETER Silent
        When specified, runs without any console output for background initialization.
    .PARAMETER IsRefresh
        When specified, changes output messages to use "refresh" terminology instead of "initialize".
    .PARAMETER CloudsToInitialize
        Specific clouds to initialize. If not provided, will initialize only missing clouds.
    .EXAMPLE
        Initialize-ExtensionCache
        
        Loads and caches missing extension data from cloud configurations.
    .EXAMPLE
        Initialize-ExtensionCache -Silent
        
        Loads missing cache silently without any output.
    .EXAMPLE
        Initialize-ExtensionCache -CloudsToInitialize @('production', 'fairfax')
        
        Loads only the specified clouds.
    .NOTES
        - Caches data in the script-scoped $ExtensionCache and $ExtensionCacheByOAuthClientId variables
        - Uses Azure DevOps CLI to fetch the file content
        - Creates nested hashtables with PortalName/OAuthClientId as key and cloud as subkey
        - Only initializes clouds that haven't been loaded yet unless specifically requested
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch] $Silent,
        
        [Parameter()]
        [switch] $IsRefresh,
        
        [Parameter()]
        [string[]] $CloudsToInitialize
    )
    
    try {
        # Initialize global caches if they don't exist
        if (-not $script:ExtensionCache) { $script:ExtensionCache = @{} }
        if (-not $script:ExtensionCacheByOAuthClientId) { $script:ExtensionCacheByOAuthClientId = @{} }
        if (-not $script:InitializedClouds) { $script:InitializedClouds = @{} }
        
        # Define all available clouds
        $allClouds = @('production', 'fairfax', 'mooncake', 'bleu', 'usnat', 'ussec', 'delos', 'dogfood')
        
        # Determine which clouds to process
        if ($CloudsToInitialize) {
            # Use specified clouds (validate they're valid)
            $cloudsToProcess = $CloudsToInitialize | Where-Object { $_ -in $allClouds }
            if ($IsRefresh) {
                # For refresh, reinitialize even if already loaded
                $finalClouds = $cloudsToProcess
            } else {
                # For normal initialization, only process clouds that aren't already initialized
                $finalClouds = $cloudsToProcess | Where-Object { -not $script:InitializedClouds.ContainsKey($_) }
            }
        } else {
            # No specific clouds provided - determine missing clouds
            if ($IsRefresh) {
                # For refresh, process all clouds
                $finalClouds = $allClouds
                # Clear existing initialization tracking for refresh
                $script:InitializedClouds = @{}
            } else {
                # For normal initialization, only process missing clouds
                $finalClouds = $allClouds | Where-Object { -not $script:InitializedClouds.ContainsKey($_) }
            }
        }
        
        # Early exit if no clouds need processing
        if ($finalClouds.Count -eq 0) {
            if (-not $Silent) {
                Write-Host "✓ All requested clouds already initialized" -ForegroundColor Green
            }
            return 0
        }
        $totalExtensions = 0
        $successfulClouds = @()
        $failedClouds = @()
        
        if (-not $Silent) {
            Write-Host "Downloading and caching extension configuration for clouds: $($finalClouds -join ', ')..." -ForegroundColor Cyan
        }
        
        # Start timing for overall operation
        $overallStartTime = Get-Date
        
        # Start all cloud fetches in parallel using background jobs
        $jobs = @()
        foreach ($cloud in $finalClouds) {
            $jobs += Start-Job -ScriptBlock {
                param($cloud, $silent)
                
                $cloudStartTime = Get-Date
                $result = @{
                    cloud = $cloud
                    success = $false
                    extensions = $null
                    extensionCount = 0
                    error = $null
                    duration = 0
                }
                
                try {
                    $errorRedirect = if ($silent) { '2>$null' } else { '' }
                    $command = "az devops invoke --organization https://dev.azure.com/msazure --area git --resource items --route-parameters project=One repositoryId=AzureUX-ExtensionStudio --query-parameters path=`"/src/roles/fusionrp/appsettings.${cloud}.json`" versionDescriptor.version=main versionDescriptor.versionType=branch includeContent=true --http-method GET --only-show-errors -o json $errorRedirect"
                    $content = Invoke-Expression $command
                    
                    if (-not $content) {
                        $result.error = "fetch failed"
                        return $result
                    }
                    
                    # Parse the response and extract the content
                    $response = $content | ConvertFrom-Json
                    $fileContent = $response.content
                    
                    if (-not $fileContent) {
                        $result.error = "no content"
                        return $result
                    }
                    
                    # Parse the JSON content
                    $appSettings = $fileContent | ConvertFrom-Json
                    
                    # Navigate to Service.HostingService.Extensions
                    $extensions = $appSettings.Service.HostingService.Extensions
                    
                    if (-not $extensions) {
                        $result.error = "no extensions"
                        return $result
                    }
                    
                    $result.success = $true
                    $result.extensions = $extensions
                    $result.extensionCount = $extensions.Count
                    
                    return $result
                }
                catch {
                    $result.error = $_.Exception.Message
                    return $result
                }
                finally {
                    $cloudEndTime = Get-Date
                    $result.duration = ($cloudEndTime - $cloudStartTime).TotalSeconds
                }
            } -ArgumentList $cloud, $Silent
        }
        
        # Show progress while jobs are running
        if (-not $Silent) {
            $spinnerChars = @('⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏')
            $spinnerIndex = 0
            
            while (($jobs | Where-Object { $_.State -eq 'Running' }).Count -gt 0) {
                $spinner = $spinnerChars[$spinnerIndex % $spinnerChars.Length]
                $completedCount = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
                Write-Host "`r$spinner Fetching $($finalClouds.Count) clouds in parallel... ($completedCount/$($finalClouds.Count) complete)" -NoNewline -ForegroundColor Yellow
                $spinnerIndex++
                Start-Sleep -Milliseconds 100
            }
            # Clear the progress line completely
            Write-Host "`r$(' ' * 80)`r" -NoNewline
        }
        
        # Collect results from all jobs
        $totalExtensions = 0
        $successfulClouds = @()
        $failedClouds = @()
        
        foreach ($job in $jobs) {
            $result = Receive-Job -Job $job -Wait
            Remove-Job -Job $job
            
            if ($result.success) {
                # Create the nested lookup caches indexed by PortalName/OAuthClientId -> Cloud -> Extension
                foreach ($extension in $result.extensions) {
                    if ($extension.PortalName) {
                        if (-not $script:ExtensionCache.ContainsKey($extension.PortalName)) {
                            $script:ExtensionCache[$extension.PortalName] = @{}
                        }
                        $script:ExtensionCache[$extension.PortalName][$result.cloud] = $extension
                    }
                    if ($extension.OAuthClientId) {
                        if (-not $script:ExtensionCacheByOAuthClientId.ContainsKey($extension.OAuthClientId)) {
                            $script:ExtensionCacheByOAuthClientId[$extension.OAuthClientId] = @{}
                        }
                        if (-not $script:ExtensionCacheByOAuthClientId[$extension.OAuthClientId].ContainsKey($result.cloud)) {
                            $script:ExtensionCacheByOAuthClientId[$extension.OAuthClientId][$result.cloud] = @()
                        }
                        $script:ExtensionCacheByOAuthClientId[$extension.OAuthClientId][$result.cloud] += $extension
                    }
                }
                
                # Mark this cloud as initialized
                $script:InitializedClouds[$result.cloud] = $true
                $successfulClouds += $result.cloud
                $totalExtensions += $result.extensionCount
                
                # Show success for each cloud
                if (-not $Silent) {
                    Write-Host "✅ $($result.cloud) - $($result.extensionCount) extensions" -ForegroundColor Green
                }
            } else {
                $failedClouds += $result.cloud
                if (-not $Silent) {
                    Write-Host "⚠️  $($result.cloud) cloud - $($result.error)" -ForegroundColor Yellow
                }
            }
        }
        
        # Clear any remaining spinner and show final results
        if (-not $Silent) {
            $overallEndTime = Get-Date
            $overallDuration = ($overallEndTime - $overallStartTime).TotalSeconds
            $actionWord = if ($IsRefresh) { "refresh" } else { "initialization" }
            Write-Host "✅ Extension cache $actionWord complete! [Total: $([math]::Round($overallDuration, 1))s]" -ForegroundColor Green
            Write-Host "   ☁️  Successful clouds: $($successfulClouds.Count)/$($finalClouds.Count) ($($successfulClouds -join ', '))" -ForegroundColor Cyan
            if ($failedClouds.Count -gt 0) {
                Write-Host "   ⚠️  Failed clouds: $($failedClouds -join ', ')" -ForegroundColor Yellow
            }
            Write-Host ""
        }
        
        return $totalExtensions
    }
    catch {
        if (-not $Silent) {
            $actionWord = if ($IsRefresh) { "refresh" } else { "initialize" }
            Write-Host "❌ Failed to $actionWord extension cache: $($_.Exception.Message)" -ForegroundColor Red
        }
        throw
    }
}

function Get-ExtensionDetails {
    <#
    .SYNOPSIS
        Interactive extension configuration lookup tool.
    .DESCRIPTION
        Prompts the user for extension names or OAuthClientIds and cloud environment for each search.
        Performs lookups against both the Extension Name cache and OAuthClientId cache across different 
        cloud environments. Continues prompting until user presses enter to exit.
    .PARAMETER RefreshCache
        Forces a refresh of the cache before lookup.
    .EXAMPLE
        Get-ExtensionDetails
        
        Starts the interactive lookup session.
    .NOTES
        - Automatically initializes cache on first use for each cloud
        - Searches both Extension Name and OAuthClientId caches for each input
        - Press enter to exit the interactive session
        - Displays full extension JSON when found
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch] $RefreshCache
    )
    
    try {
        Write-Host "=== Interactive Extension Lookup ===" -ForegroundColor Cyan
        Write-Host "Enter extension name or OAuthClientId. Press ENTER to exit." -ForegroundColor Gray
        Write-Host ""
        
        do {
            # Prompt for search term
            $searchTerm = Read-Host "Search (extension name or OAuthClientId)"
            
            # Check if user wants to exit (empty)
            if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                Write-Host "Exiting extension lookup." -ForegroundColor Gray
                break
            }
            
            # Prompt for cloud with validation
            $validClouds = @('production', 'fairfax', 'mooncake', 'bleu', 'usnat', 'ussec', 'delos', 'dogfood')
            $cloudAcronyms = @{
                'prod' = 'production'
                'ff' = 'fairfax'
                'mc' = 'mooncake'
                'df' = 'dogfood'
                'usnat' = 'usnat'
                'ussec' = 'ussec'
                'bleu' = 'bleu'
                'delos' = 'delos'
            }
            do {
                $cloud = Read-Host "Cloud: (prod(default)/ff/mc/bleu/usnat/ussec/delos/df)"
                if ([string]::IsNullOrWhiteSpace($cloud)) {
                    $cloud = "production"
                    break
                }
                $cloud = $cloud.Trim().ToLower()
                
                # Check if it's an acronym and convert it
                if ($cloudAcronyms.ContainsKey($cloud)) {
                    $cloud = $cloudAcronyms[$cloud]
                }
                
                if ($cloud -notin $validClouds) {
                    Write-Host "Invalid cloud. Please enter one of: $($validClouds -join ', ') or their acronyms: $($cloudAcronyms.Keys -join ', ')" -ForegroundColor Red
                    $cloud = $null
                }
            } while ([string]::IsNullOrWhiteSpace($cloud))
            
            # Initialize cache if not loaded or if refresh requested
            if (-not $script:ExtensionCache -or -not $script:ExtensionCacheByOAuthClientId -or $RefreshCache) {
                if ($RefreshCache) {
                    Initialize-ExtensionCache -IsRefresh | Out-Null
                } else {
                    # Check if we need to initialize any clouds
                    $allClouds = @('production', 'fairfax', 'mooncake', 'bleu', 'usnat', 'ussec', 'delos', 'dogfood')
                    $missingClouds = $allClouds | Where-Object { -not $script:InitializedClouds.ContainsKey($_) }
                    
                    if ($missingClouds.Count -gt 0) {
                        Initialize-ExtensionCache -CloudsToInitialize $missingClouds | Out-Null
                    }
                }
            }
            
            $searchTerm = $searchTerm.Trim()
            $extension = $null
            $extensions = @()
            
            # First try Extension Name lookup
            if ($script:ExtensionCache.ContainsKey($searchTerm) -and $script:ExtensionCache[$searchTerm].ContainsKey($cloud)) {
                $extension = $script:ExtensionCache[$searchTerm][$cloud]
            }
            # Then try OAuthClientId lookup
            elseif ($script:ExtensionCacheByOAuthClientId.ContainsKey($searchTerm) -and $script:ExtensionCacheByOAuthClientId[$searchTerm].ContainsKey($cloud)) {
                $extensions = $script:ExtensionCacheByOAuthClientId[$searchTerm][$cloud]
                if ($extensions.Count -eq 1) {
                    $extension = $extensions[0]
                    $extensions = @()
                }
            }
            
            if ($extension) {
                # Display single extension details
                Write-Host "=== Extension Details ($cloud cloud) ===" -ForegroundColor Cyan
                
                # Convert extension to JSON first
                $jsonOutput = $extension | ConvertTo-Json -Depth 10
                
                # Add hyperlinks by replacing specific patterns in the JSON string
                if ($extension.Icm.TeamId) {
                    $teamId = $extension.Icm.TeamId
                    $onCallLink = "https://portal.microsofticm.com/imp/v3/oncall/current?teamIds=${teamId}&scheduleType=current&shiftType=current&viewType=1"
                    $onCallHyperlink = "`e]8;;${onCallLink}`e\`e[94mOn-call link`e[0m`e]8;;`e\"
                    $jsonOutput = $jsonOutput -replace "(`"TeamId`":\s*)($teamId)", "`$1`$2 ($onCallHyperlink)"
                }
                
                if ($extension.OAuthClientId) {
                    $appID = $extension.OAuthClientId
                    $appConfigLink = "https://msazure.visualstudio.com/One/_git/AAD-FirstPartyApps?path=/Customers/Configs/AppReg/${appID}"
                    $appConfigHyperlink = "`e]8;;${appConfigLink}`e\`e[94m1P app config link`e[0m`e]8;;`e\"
                    $jsonOutput = $jsonOutput -replace "(`"OAuthClientId`":\s*`")($appID)(`")", "`$1`$2 ($appConfigHyperlink)`$3"
                }
                
                # Highlight the search term, but avoid highlighting within hyperlink escape sequences
                $escapedSearchTerm = [regex]::Escape($searchTerm)
                $highlightedSearchTerm = "`e[93m$searchTerm`e[0m"  # Yellow highlight
                # Use negative lookahead/lookbehind to avoid highlighting within ANSI escape sequences
                $jsonOutput = $jsonOutput -replace "(?<!`e\]8;;[^`e]*?)($escapedSearchTerm)(?![^`e]*?`e\\)", $highlightedSearchTerm
                
                # Display the JSON with embedded hyperlinks and highlighted search terms
                Write-Host $jsonOutput -ForegroundColor White
                
                # Show if extension exists in other clouds
                $otherClouds = @()
                if ($script:ExtensionCache.ContainsKey($searchTerm)) {
                    $otherClouds += $script:ExtensionCache[$searchTerm].Keys | Where-Object { $_ -ne $cloud }
                }
                elseif ($script:ExtensionCacheByOAuthClientId.ContainsKey($searchTerm)) {
                    $otherClouds += $script:ExtensionCacheByOAuthClientId[$searchTerm].Keys | Where-Object { $_ -ne $cloud }
                }
                
                if ($otherClouds.Count -gt 0) {
                    Write-Host "💡 This extension also exists in: $($otherClouds -join ', ')" -ForegroundColor Cyan
                    Write-Host "=== End ===" -ForegroundColor Cyan
                }
            }
            elseif ($extensions.Count -gt 1) {
                # Multiple extensions found for this OAuthClientId - display list and return to parent search
                Write-Host "=== Multiple Extensions Found for OAuthClientId '$searchTerm' in $cloud cloud ===" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "Found $($extensions.Count) extensions using this OAuthClientId:" -ForegroundColor White
                Write-Host ""
                
                for ($i = 0; $i -lt $extensions.Count; $i++) {
                    $ext = $extensions[$i]
                    Write-Host "  $($i + 1). " -NoNewline -ForegroundColor Yellow
                    Write-Host $ext.PortalName -ForegroundColor White
                }
                Write-Host ""
                Write-Host "💡 Search by extension name to view specific details" -ForegroundColor Cyan
                Write-Host "=== End ===" -ForegroundColor Cyan
            }
            else {
                Write-Host "⚠ Extension '$searchTerm' not found as exact match in $cloud cloud" -ForegroundColor Yellow
                
                # Find similar extensions using substring and left-match criteria
                $suggestions = @()
                $searchLower = $searchTerm.ToLower()
                
                # Get all unique extension names from both caches that exist in this cloud
                $allExtensionNames = @()
                
                # Add extension names from ExtensionCache
                if ($script:ExtensionCache -and $script:ExtensionCache.Count -gt 0) {
                    foreach ($extensionName in $script:ExtensionCache.Keys) {
                        try {
                            if ($script:ExtensionCache[$extensionName].ContainsKey($cloud) -and 
                                $extensionName -and
                                $extensionName.GetType() -eq [string] -and
                                $extensionName.Length -gt 1) {
                                $allExtensionNames += $extensionName
                            }
                        }
                        catch {
                            # Skip problematic entries
                            continue
                        }
                    }
                }
                
                # Add portal names from OAuthClientId cache
                if ($script:ExtensionCacheByOAuthClientId -and $script:ExtensionCacheByOAuthClientId.Count -gt 0) {
                    foreach ($oauthId in $script:ExtensionCacheByOAuthClientId.Keys) {
                        try {
                            if ($script:ExtensionCacheByOAuthClientId[$oauthId].ContainsKey($cloud)) {
                                $extensionsForOAuth = $script:ExtensionCacheByOAuthClientId[$oauthId][$cloud]
                                
                                # Handle both single extension and array of extensions
                                $extensionList = if ($extensionsForOAuth -is [array]) { $extensionsForOAuth } else { @($extensionsForOAuth) }
                                
                                foreach ($ext in $extensionList) {
                                    if ($ext -and $ext.PortalName -and $ext.PortalName.GetType() -eq [string] -and $ext.PortalName.Length -gt 1) {
                                        $allExtensionNames += $ext.PortalName
                                    }
                                }
                            }
                        }
                        catch {
                            # Skip problematic entries
                            continue
                        }
                    }
                }
                
                # Remove duplicates and filter out invalid entries
                $allExtensionNames = $allExtensionNames | Where-Object { 
                    $_ -and $_.GetType() -eq [string] -and $_.Length -gt 1 
                } | Select-Object -Unique | Sort-Object
                
                if ($allExtensionNames.Count -gt 0) {
                    # Find extensions that start with the search term (left match) - higher priority
                    $leftMatches = $allExtensionNames | Where-Object { 
                        try { $_.ToLower().StartsWith($searchLower) } catch { $false }
                    }
                    
                    # Find extensions that contain the search term as substring
                    $substringMatches = $allExtensionNames | Where-Object { 
                        try { $_.ToLower().Contains($searchLower) -and -not $_.ToLower().StartsWith($searchLower) } catch { $false }
                    }
                    
                    # Combine results, prioritizing left matches, then limit to 10 total
                    $suggestions = @()
                    if ($leftMatches -and $leftMatches.Count -gt 0) { 
                        $suggestions += @($leftMatches)
                    }
                    if ($substringMatches -and $substringMatches.Count -gt 0) { 
                        $suggestions += @($substringMatches)
                    }
                    
                    # Ensure we have a proper array and limit to 10 total
                    if ($suggestions.Count -gt 0) {
                        $suggestions = @($suggestions | Select-Object -First 10)
                    }
                }
                
                if ($suggestions.Count -gt 0) {
                    Write-Host ""
                    Write-Host "Did you mean one of these (in $cloud cloud)?" -ForegroundColor Cyan
                    Write-Host ""
                    
                    for ($i = 0; $i -lt $suggestions.Count; $i++) {
                        $suggestion = $suggestions[$i]
                        Write-Host "  $($i + 1). " -NoNewline -ForegroundColor Gray
                        Write-Host $suggestion -ForegroundColor White
                    }
                }
                else {
                    Write-Host "No similar extension names found in $cloud cloud." -ForegroundColor Gray
                }
            }
            
            Write-Host ""
            
        } while ($true)
    }
    catch {
        Write-Host "❌ Failed to retrieve extension data: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

#==============================================================================
# 1P APP CONFIGURATION SCANNING
#==============================================================================
#
# This section provides tools to scan 1P (First Party) app configurations across
# all extensions in a given cloud environment. The scanning system allows users to
# define criteria (exact match or CONTAINS) and evaluates each extension's 1P app
# config against that criteria.
#
# KEY FEATURES:
# - Downloads AppReg.Parameters.<cloud>.json from the AAD-FirstPartyApps repo
# - Supports exact string match and CONTAINS array checks
# - Deduplicates app ID downloads (many extensions share the same OAuthClientId)
# - Parallel downloading of 1P app configs for performance
# - Verbose diagnostic logging for troubleshooting
# - Reports extensions that do NOT satisfy the user-defined criteria
#
# URL FORMAT:
#   https://msazure.visualstudio.com/One/_git/AAD-FirstPartyApps
#     ?path=/Customers/Configs/AppReg/{appID}/AppReg.Parameters.{cloud}.json
#
# CRITERIA FORMAT:
#   Exact match:  parameters.signInAudience.value == "AzureADMultipleOrgs"
#   Contains:     parameters.spa.value.redirectUris CONTAINS ["https://url1", "https://url2"]
#
#==============================================================================

# Mapping from internal cloud names to the cloud name used in AppReg.Parameters.<cloud>.json file paths
$script:CloudToAppRegFileNameMap = @{
    'production' = 'Prod'
    'fairfax'    = 'Fairfax'
    'mooncake'   = 'Mooncake'
    'bleu'       = 'Bleu'
    'usnat'      = 'USNat'
    'ussec'      = 'USSec'
    'delos'      = 'Delos'
    'dogfood'    = 'Dogfood'
}

function Get-FirstPartyAppConfig {
    <#
    .SYNOPSIS
        Downloads the 1P app configuration for a given app ID and cloud from the AAD-FirstPartyApps repo.
    .DESCRIPTION
        Uses the Azure DevOps git API to fetch AppReg.Parameters.<cloud>.json for a specific app ID
        from the AAD-FirstPartyApps repository. Returns the parsed JSON object or $null on failure.
    .PARAMETER AppId
        The OAuthClientId / App ID to look up.
    .PARAMETER CloudFileName
        The cloud file name suffix (e.g. 'Prod', 'Fairfax') used in AppReg.Parameters.<cloud>.json.
    .OUTPUTS
        PSCustomObject representing the parsed 1P app config, or $null if not found/error.
    .EXAMPLE
        $config = Get-FirstPartyAppConfig -AppId "00000000-0000-0000-0000-000000000000" -CloudFileName "Prod"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $AppId,

        [Parameter(Mandatory)]
        [string] $CloudFileName
    )

    $filePath = "/Customers/Configs/AppReg/$AppId/AppReg.Parameters.$CloudFileName.json"
    Write-Host "    [FETCH] Downloading 1P app config: repo=AAD-FirstPartyApps, path=$filePath" -ForegroundColor DarkGray

    try {
        $rawOutput = az devops invoke `
            --organization https://dev.azure.com/msazure `
            --area git --resource items `
            --route-parameters project=One repositoryId=AAD-FirstPartyApps `
            --query-parameters "path=$filePath" versionDescriptor.version=master versionDescriptor.versionType=branch includeContent=true `
            --http-method GET --only-show-errors -o json 2>&1

        if ($LASTEXITCODE -ne 0) {
            Write-Host "    [FETCH] az devops invoke failed with exit code $LASTEXITCODE for AppId=$AppId" -ForegroundColor DarkGray
            Write-Host "    [FETCH] Output: $rawOutput" -ForegroundColor DarkGray
            return $null
        }

        $response = $rawOutput | ConvertFrom-Json
        $fileContent = $response.content

        if (-not $fileContent) {
            Write-Host "    [FETCH] No content in response for AppId=$AppId" -ForegroundColor DarkGray
            return $null
        }

        $appConfig = $fileContent | ConvertFrom-Json
        Write-Host "    [FETCH] Successfully parsed 1P app config for AppId=$AppId" -ForegroundColor DarkGray
        return $appConfig
    }
    catch {
        Write-Host "    [FETCH] Error downloading 1P app config for AppId=$AppId : $($_.Exception.Message)" -ForegroundColor DarkGray
        return $null
    }
}

function Resolve-JsonPropertyPath {
    <#
    .SYNOPSIS
        Navigates a JSON/PSCustomObject by a dot-separated property path.
    .DESCRIPTION
        Traverses nested properties of a PowerShell object using a dot-separated path string.
        Supports array indexing with bracket notation (e.g. "items[0].name").
        Returns $null if any segment of the path does not exist.
    .PARAMETER Object
        The root object to navigate.
    .PARAMETER PropertyPath
        Dot-separated property path (e.g. "parameters.signInAudience.value").
    .OUTPUTS
        The value at the specified path, or $null if not found.
    .EXAMPLE
        $value = Resolve-JsonPropertyPath -Object $config -PropertyPath "parameters.signInAudience.value"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Object,

        [Parameter(Mandatory)]
        [string] $PropertyPath
    )

    $current = $Object
    $segments = $PropertyPath.Split('.')

    Write-Host "      [PATH] Resolving property path '$PropertyPath' ($($segments.Count) segments)" -ForegroundColor DarkGray

    foreach ($segment in $segments) {
        if ($null -eq $current) {
            Write-Host "      [PATH] Hit null at segment '$segment'" -ForegroundColor DarkGray
            return $null
        }

        # Handle array indexing like property[0]
        if ($segment -match '^(.+)\[(\d+)\]$') {
            $propName = $Matches[1]
            $index = [int]$Matches[2]

            Write-Host "      [PATH] Navigating '$propName' then index [$index]" -ForegroundColor DarkGray

            $current = $current.$propName
            if ($null -eq $current) {
                Write-Host "      [PATH] Property '$propName' is null" -ForegroundColor DarkGray
                return $null
            }
            $arr = @($current)
            if ($index -ge $arr.Count) {
                Write-Host "      [PATH] Index $index out of range (array length: $($arr.Count))" -ForegroundColor DarkGray
                return $null
            }
            $current = $arr[$index]
        }
        else {
            $prev = $current
            $current = $current.$segment
            if ($null -eq $current) {
                # Check if the property actually exists but has a null value vs not existing
                $propExists = ($prev.PSObject.Properties.Name -contains $segment)
                if ($propExists) {
                    Write-Host "      [PATH] Property '$segment' exists but has null value" -ForegroundColor DarkGray
                } else {
                    Write-Host "      [PATH] Property '$segment' does NOT exist on object" -ForegroundColor DarkGray
                    Write-Host "      [PATH] Available properties: $($prev.PSObject.Properties.Name -join ', ')" -ForegroundColor DarkGray
                }
                return $null
            }
            Write-Host "      [PATH] '$segment' => type=$($current.GetType().Name)" -ForegroundColor DarkGray
        }
    }

    return $current
}

function Test-AppConfigCriteria {
    <#
    .SYNOPSIS
        Evaluates a criteria expression against a 1P app configuration object.
    .DESCRIPTION
        Supports two criteria types:
        - 'exact': checks if the property at the given path equals the expected string value
        - 'contains': checks if the array property at the given path contains EACH of the expected values independently
        Returns a hashtable with per-value results so that partial matches can be reported.
    .PARAMETER AppConfig
        The parsed 1P app configuration PSCustomObject.
    .PARAMETER CriteriaType
        Either 'exact' or 'contains'.
    .PARAMETER PropertyPath
        Dot-separated property path to evaluate (e.g. "parameters.signInAudience.value").
    .PARAMETER ExpectedValue
        For 'exact': a string to compare against.
        For 'contains': an array of strings that must all be present in the array at the property path.
    .OUTPUTS
        Hashtable with keys:
          - AllSatisfied: $true/$false — whether every value was found
          - PerValue: ordered hashtable mapping each expected value to $true/$false
    .EXAMPLE
        $result = Test-AppConfigCriteria -AppConfig $cfg -CriteriaType 'exact' -PropertyPath 'parameters.signInAudience.value' -ExpectedValue 'AzureADMultipleOrgs'
        $result.AllSatisfied  # $true or $false
    .EXAMPLE
        $result = Test-AppConfigCriteria -AppConfig $cfg -CriteriaType 'contains' -PropertyPath 'parameters.spa.value.redirectUris' -ExpectedValue @("https://url1", "https://url2")
        $result.PerValue      # ordered hashtable: url1 -> $true, url2 -> $false
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $AppConfig,

        [Parameter(Mandatory)]
        [ValidateSet('exact', 'contains')]
        [string] $CriteriaType,

        [Parameter(Mandatory)]
        [string] $PropertyPath,

        [Parameter(Mandatory)]
        $ExpectedValue
    )

    Write-Host "    [EVAL] Evaluating criteria: type=$CriteriaType, path='$PropertyPath'" -ForegroundColor DarkGray

    $actualValue = Resolve-JsonPropertyPath -Object $AppConfig -PropertyPath $PropertyPath

    # Build the per-value results hashtable
    $perValue = [ordered]@{}

    if ($null -eq $actualValue) {
        Write-Host "    [EVAL] Property path '$PropertyPath' resolved to null => all criteria NOT satisfied" -ForegroundColor DarkGray
        if ($CriteriaType -eq 'exact') {
            $perValue[$ExpectedValue] = $false
        } else {
            foreach ($ev in $ExpectedValue) { $perValue[$ev.Trim()] = $false }
        }
        return @{ AllSatisfied = $false; PerValue = $perValue }
    }

    if ($CriteriaType -eq 'exact') {
        $actualStr = $actualValue.ToString()
        $matched = ($actualStr -eq $ExpectedValue)
        Write-Host "    [EVAL] Exact match: actual='$actualStr' == expected='$ExpectedValue' => $matched" -ForegroundColor DarkGray
        $perValue[$ExpectedValue] = $matched
        return @{ AllSatisfied = $matched; PerValue = $perValue }
    }
    elseif ($CriteriaType -eq 'contains') {
        # actualValue should be an array; check each expected value independently
        $actualArray = @($actualValue)
        Write-Host "    [EVAL] CONTAINS check: actual array has $($actualArray.Count) items, checking for $($ExpectedValue.Count) expected values" -ForegroundColor DarkGray

        $allFound = $true
        foreach ($expected in $ExpectedValue) {
            $trimmedExpected = $expected.Trim()
            $found = $false
            foreach ($item in $actualArray) {
                if ($item.ToString().Trim() -eq $trimmedExpected) {
                    $found = $true
                    break
                }
            }
            $perValue[$trimmedExpected] = $found
            if ($found) {
                Write-Host "    [EVAL]   ✓ Found: '$trimmedExpected'" -ForegroundColor DarkGray
            } else {
                Write-Host "    [EVAL]   ✗ NOT Found: '$trimmedExpected'" -ForegroundColor DarkGray
                $allFound = $false
            }
        }

        Write-Host "    [EVAL] CONTAINS result: AllSatisfied=$allFound" -ForegroundColor DarkGray
        return @{ AllSatisfied = $allFound; PerValue = $perValue }
    }

    Write-Host "    [EVAL] Unknown criteria type '$CriteriaType' => false" -ForegroundColor DarkGray
    return @{ AllSatisfied = $false; PerValue = $perValue }
}

function Parse-ScanCriteria {
    <#
    .SYNOPSIS
        Parses a user-provided criteria string into its components.
    .DESCRIPTION
        Supports two formats:
        - Exact match:  propertyPath == "value"
        - Contains:     propertyPath CONTAINS ["value1", "value2", ...]
        Returns a hashtable with CriteriaType, PropertyPath, and ExpectedValue.
    .PARAMETER CriteriaString
        The raw criteria string entered by the user.
    .OUTPUTS
        Hashtable with keys: CriteriaType ('exact'|'contains'), PropertyPath (string), ExpectedValue (string or string[]).
        Returns $null if parsing fails.
    .EXAMPLE
        Parse-ScanCriteria 'parameters.signInAudience.value == "AzureADMultipleOrgs"'
    .EXAMPLE
        Parse-ScanCriteria 'parameters.spa.value.redirectUris CONTAINS ["https://url1", "https://url2"]'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $CriteriaString
    )

    $criteria = $CriteriaString.Trim()
    Write-Host "[PARSE] Parsing criteria: '$criteria'" -ForegroundColor DarkGray

    # Try exact match format: propertyPath == "value"
    if ($criteria -match '^\s*(.+?)\s*==\s*"(.*)"\s*$') {
        $propertyPath = $Matches[1].Trim()
        $expectedValue = $Matches[2]
        Write-Host "[PARSE] Detected EXACT match criteria" -ForegroundColor DarkGray
        Write-Host "[PARSE]   PropertyPath : $propertyPath" -ForegroundColor DarkGray
        Write-Host "[PARSE]   ExpectedValue: $expectedValue" -ForegroundColor DarkGray
        return @{
            CriteriaType  = 'exact'
            PropertyPath  = $propertyPath
            ExpectedValue = $expectedValue
        }
    }

    # Try contains format: propertyPath CONTAINS [...]
    if ($criteria -match '^\s*(.+?)\s+CONTAINS\s+(\[[\s\S]*\])\s*$') {
        $propertyPath = $Matches[1].Trim()
        $jsonArrayStr = $Matches[2].Trim()
        Write-Host "[PARSE] Detected CONTAINS criteria" -ForegroundColor DarkGray
        Write-Host "[PARSE]   PropertyPath  : $propertyPath" -ForegroundColor DarkGray
        Write-Host "[PARSE]   Raw JSON array: $jsonArrayStr" -ForegroundColor DarkGray

        try {
            $parsedArray = $jsonArrayStr | ConvertFrom-Json
            $expectedValues = @($parsedArray)
            Write-Host "[PARSE]   Parsed $($expectedValues.Count) expected values" -ForegroundColor DarkGray
            foreach ($val in $expectedValues) {
                Write-Host "[PARSE]     - '$val'" -ForegroundColor DarkGray
            }
            return @{
                CriteriaType  = 'contains'
                PropertyPath  = $propertyPath
                ExpectedValue = $expectedValues
            }
        }
        catch {
            Write-Host "[PARSE] ❌ Failed to parse JSON array: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "[PARSE]   Raw input was: $jsonArrayStr" -ForegroundColor Red
            return $null
        }
    }

    Write-Host "[PARSE] ❌ Could not parse criteria. Expected format:" -ForegroundColor Red
    Write-Host "[PARSE]   Exact:    propertyPath == `"value`"" -ForegroundColor Yellow
    Write-Host "[PARSE]   Contains: propertyPath CONTAINS [`"value1`", `"value2`"]" -ForegroundColor Yellow
    return $null
}

function Invoke-FirstPartyAppScan {
    <#
    .SYNOPSIS
        Scans 1P app configurations for all extensions in a cloud against user-defined criteria.
    .DESCRIPTION
        This function iterates through all extensions registered in a cloud-specific
        appsettings.<cloud>.json file. For each extension that has an OAuthClientId,
        it downloads the corresponding 1P app config (AppReg.Parameters.<cloud>.json)
        from the AAD-FirstPartyApps repo and evaluates the user-provided criteria.

        The function reports all extensions whose 1P app configs do NOT satisfy the criteria.

        Two criteria formats are supported:
        - Exact match:  propertyPath == "value"
          e.g. parameters.signInAudience.value == "AzureADMultipleOrgs"

        - Contains:     propertyPath CONTAINS ["value1", "value2"]
          e.g. parameters.spa.value.redirectUris CONTAINS ["https://url1", "https://url2"]
    .PARAMETER Cloud
        The cloud to scan. If not provided, the user will be prompted.
    .PARAMETER Criteria
        The criteria expression string. If not provided, the user will be prompted.
    .PARAMETER RefreshCache
        Forces a refresh of the extension cache before scanning.
    .PARAMETER BatchSize
        Number of parallel app config downloads at a time. Defaults to 10.
    .EXAMPLE
        Invoke-FirstPartyAppScan
        # Interactive mode - prompts for cloud and criteria

    .EXAMPLE
        Invoke-FirstPartyAppScan -Cloud 'fairfax' -Criteria 'parameters.signInAudience.value == "AzureADMultipleOrgs"'
        # Non-interactive mode with parameters

    .EXAMPLE
        scanapps
        # Using the short alias
    .NOTES
        - Requires the extension cache to be initialized (will auto-initialize if needed)
        - Downloads are batched in parallel for performance
        - Deduplicates app IDs (many extensions share the same OAuthClientId)
        - Extensions without an OAuthClientId are reported separately
        - Uses extensive logging for debugging (all lines prefixed with [tags])
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string] $Cloud,

        [Parameter()]
        [string] $Criteria,

        [Parameter()]
        [switch] $RefreshCache,

        [Parameter()]
        [int] $BatchSize = 10
    )

    try {
        Write-Host ""
        Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║      1P App Configuration Scanner               ║" -ForegroundColor Cyan
        Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""

        # -----------------------------------------------------------------
        # STEP 1: Determine cloud
        # -----------------------------------------------------------------
        $validClouds = @('production', 'fairfax', 'mooncake', 'bleu', 'usnat', 'ussec', 'delos', 'dogfood')
        $cloudAcronyms = @{
            'prod' = 'production'; 'ff' = 'fairfax'; 'mc' = 'mooncake'
            'df'   = 'dogfood';    'usnat' = 'usnat'; 'ussec' = 'ussec'
            'bleu' = 'bleu';       'delos' = 'delos'
        }

        if (-not $Cloud) {
            do {
                $Cloud = Read-Host "Cloud to scan (prod(default)/ff/mc/bleu/usnat/ussec/delos/df)"
                if ([string]::IsNullOrWhiteSpace($Cloud)) {
                    $Cloud = 'production'
                    break
                }
                $Cloud = $Cloud.Trim().ToLower()
                if ($cloudAcronyms.ContainsKey($Cloud)) { $Cloud = $cloudAcronyms[$Cloud] }
                if ($Cloud -notin $validClouds) {
                    Write-Host "Invalid cloud. Valid: $($validClouds -join ', ')" -ForegroundColor Red
                    $Cloud = $null
                }
            } while ([string]::IsNullOrWhiteSpace($Cloud))
        } else {
            $Cloud = $Cloud.Trim().ToLower()
            if ($cloudAcronyms.ContainsKey($Cloud)) { $Cloud = $cloudAcronyms[$Cloud] }
            if ($Cloud -notin $validClouds) {
                Write-Host "❌ Invalid cloud '$Cloud'. Valid: $($validClouds -join ', ')" -ForegroundColor Red
                return
            }
        }

        Write-Host "[INFO] Selected cloud: $Cloud" -ForegroundColor Cyan

        # Resolve the cloud file name for the AppReg URL
        $cloudFileName = $script:CloudToAppRegFileNameMap[$Cloud]
        if (-not $cloudFileName) {
            Write-Host "❌ No AppReg file name mapping found for cloud '$Cloud'. Check `$script:CloudToAppRegFileNameMap." -ForegroundColor Red
            return
        }
        Write-Host "[INFO] AppReg file name suffix: $cloudFileName (AppReg.Parameters.$cloudFileName.json)" -ForegroundColor Cyan

        # -----------------------------------------------------------------
        # STEP 2: Get criteria from user
        # -----------------------------------------------------------------
        if (-not $Criteria) {
            Write-Host ""
            Write-Host "Enter the criteria to evaluate against each extension's 1P app config." -ForegroundColor White
            Write-Host "Supported formats:" -ForegroundColor Gray
            Write-Host "  Exact:    propertyPath == `"value`"" -ForegroundColor Gray
            Write-Host "  Contains: propertyPath CONTAINS [`"value1`", `"value2`"]" -ForegroundColor Gray
            Write-Host ""
            Write-Host "Examples:" -ForegroundColor Gray
            Write-Host "  parameters.signInAudience.value == `"AzureADMultipleOrgs`"" -ForegroundColor DarkYellow
            Write-Host "  parameters.spa.value.redirectUris CONTAINS [`"https://canary.entra.microsoft.eaglex.ic.gov/auth/login/`"]" -ForegroundColor DarkYellow
            Write-Host ""
            $Criteria = Read-Host "Criteria"
        }

        if ([string]::IsNullOrWhiteSpace($Criteria)) {
            Write-Host "❌ No criteria provided. Aborting." -ForegroundColor Red
            return
        }

        # Parse the criteria
        $parsedCriteria = Parse-ScanCriteria -CriteriaString $Criteria
        if (-not $parsedCriteria) {
            Write-Host "❌ Failed to parse criteria. Aborting." -ForegroundColor Red
            return
        }

        Write-Host "[INFO] Criteria parsed successfully: type=$($parsedCriteria.CriteriaType), path=$($parsedCriteria.PropertyPath)" -ForegroundColor Cyan
        Write-Host ""

        # -----------------------------------------------------------------
        # STEP 3: Ensure extension cache is initialized for the selected cloud
        # -----------------------------------------------------------------
        Write-Host "[INFO] Ensuring extension cache is initialized for '$Cloud'..." -ForegroundColor Cyan

        if ($RefreshCache) {
            Initialize-ExtensionCache -IsRefresh -CloudsToInitialize @($Cloud) | Out-Null
        } elseif (-not $script:InitializedClouds.ContainsKey($Cloud)) {
            Initialize-ExtensionCache -CloudsToInitialize @($Cloud) | Out-Null
        } else {
            Write-Host "[INFO] Extension cache already initialized for '$Cloud'" -ForegroundColor Cyan
        }

        if (-not $script:InitializedClouds.ContainsKey($Cloud)) {
            Write-Host "❌ Failed to initialize extension cache for '$Cloud'. Cannot proceed." -ForegroundColor Red
            return
        }

        # -----------------------------------------------------------------
        # STEP 4: Collect all extensions for this cloud and their unique app IDs
        # -----------------------------------------------------------------
        Write-Host "[INFO] Collecting extensions from cache for '$Cloud'..." -ForegroundColor Cyan

        $extensionsInCloud = @()
        $extensionsWithoutAppId = @()
        $uniqueAppIds = @{}   # AppId -> @() list of extension names using it

        foreach ($portalName in $script:ExtensionCache.Keys) {
            if ($script:ExtensionCache[$portalName].ContainsKey($Cloud)) {
                $ext = $script:ExtensionCache[$portalName][$Cloud]
                $extensionsInCloud += $ext

                if ($ext.OAuthClientId) {
                    $appId = $ext.OAuthClientId
                    if (-not $uniqueAppIds.ContainsKey($appId)) {
                        $uniqueAppIds[$appId] = @()
                    }
                    $uniqueAppIds[$appId] += $portalName
                } else {
                    $extensionsWithoutAppId += $portalName
                }
            }
        }

        Write-Host "[INFO] Found $($extensionsInCloud.Count) extensions in '$Cloud'" -ForegroundColor Cyan
        Write-Host "[INFO] $($uniqueAppIds.Count) unique OAuthClientIds to scan" -ForegroundColor Cyan
        Write-Host "[INFO] $($extensionsWithoutAppId.Count) extensions have no OAuthClientId" -ForegroundColor Cyan
        Write-Host ""

        if ($uniqueAppIds.Count -eq 0) {
            Write-Host "⚠ No extensions with OAuthClientId found in '$Cloud'. Nothing to scan." -ForegroundColor Yellow
            return
        }

        # -----------------------------------------------------------------
        # STEP 5: Download 1P app configs in parallel batches
        # -----------------------------------------------------------------
        Write-Host "[INFO] Downloading 1P app configs (batch size: $BatchSize)..." -ForegroundColor Cyan
        $appConfigs = @{}     # AppId -> parsed config or $null
        $appIdList = @($uniqueAppIds.Keys)
        $totalAppIds = $appIdList.Count
        $downloadedCount = 0
        $downloadFailedCount = 0

        $overallStartTime = Get-Date

        for ($batchStart = 0; $batchStart -lt $totalAppIds; $batchStart += $BatchSize) {
            $batchEnd = [math]::Min($batchStart + $BatchSize, $totalAppIds) - 1
            $batch = $appIdList[$batchStart..$batchEnd]
            $batchNum = [math]::Floor($batchStart / $BatchSize) + 1
            $totalBatches = [math]::Ceiling($totalAppIds / $BatchSize)

            Write-Host ""
            Write-Host "[BATCH $batchNum/$totalBatches] Downloading $($batch.Count) app configs (IDs $($batchStart + 1)-$($batchEnd + 1) of $totalAppIds)..." -ForegroundColor Yellow

            # Start parallel jobs for this batch
            $jobs = @()
            foreach ($appId in $batch) {
                Write-Host "  [JOB] Starting download for AppId=$appId (used by: $($uniqueAppIds[$appId] -join ', '))" -ForegroundColor DarkGray
                $jobs += Start-Job -ScriptBlock {
                    param($appId, $cloudFileName)

                    $result = @{
                        AppId   = $appId
                        Success = $false
                        Config  = $null
                        Error   = $null
                    }

                    $filePath = "/Customers/Configs/AppReg/$appId/AppReg.Parameters.$cloudFileName.json"

                    try {
                        $rawOutput = az devops invoke `
                            --organization https://dev.azure.com/msazure `
                            --area git --resource items `
                            --route-parameters project=One repositoryId=AAD-FirstPartyApps `
                            --query-parameters "path=$filePath" versionDescriptor.version=master versionDescriptor.versionType=branch includeContent=true `
                            --http-method GET --only-show-errors -o json 2>&1

                        if ($LASTEXITCODE -ne 0) {
                            $result.Error = "az devops invoke failed (exit code $LASTEXITCODE): $rawOutput"
                            return $result
                        }

                        $response = $rawOutput | ConvertFrom-Json
                        $fileContent = $response.content

                        if (-not $fileContent) {
                            $result.Error = "No content in response"
                            return $result
                        }

                        $result.Config  = $fileContent | ConvertFrom-Json
                        $result.Success = $true
                        return $result
                    }
                    catch {
                        $result.Error = $_.Exception.Message
                        return $result
                    }
                } -ArgumentList $appId, $cloudFileName
            }

            # Wait for batch to complete with spinner
            $spinnerChars = @('⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏')
            $spinnerIndex = 0
            while (($jobs | Where-Object { $_.State -eq 'Running' }).Count -gt 0) {
                $spinner = $spinnerChars[$spinnerIndex % $spinnerChars.Length]
                $completedInBatch = ($jobs | Where-Object { $_.State -eq 'Completed' }).Count
                Write-Host "`r  $spinner Waiting for batch... ($completedInBatch/$($batch.Count) complete)" -NoNewline -ForegroundColor Yellow
                $spinnerIndex++
                Start-Sleep -Milliseconds 100
            }
            Write-Host "`r  $(' ' * 60)`r" -NoNewline

            # Collect batch results
            foreach ($job in $jobs) {
                $result = Receive-Job -Job $job -Wait
                Remove-Job -Job $job

                $appId = $result.AppId
                if ($result.Success) {
                    $appConfigs[$appId] = $result.Config
                    $downloadedCount++
                    Write-Host "  [OK] AppId=$appId - config downloaded" -ForegroundColor Green
                } else {
                    $appConfigs[$appId] = $null
                    $downloadFailedCount++
                    Write-Host "  [FAIL] AppId=$appId - $($result.Error)" -ForegroundColor Red
                }
            }
        }

        $overallEndTime = Get-Date
        $totalDuration = ($overallEndTime - $overallStartTime).TotalSeconds

        Write-Host ""
        Write-Host "[INFO] Download complete in $([math]::Round($totalDuration, 1))s — Success: $downloadedCount, Failed: $downloadFailedCount, Total: $totalAppIds" -ForegroundColor Cyan
        Write-Host ""

        # -----------------------------------------------------------------
        # STEP 6: Ensure ImportExcel module is available
        # -----------------------------------------------------------------
        Write-Host "[INFO] Checking for ImportExcel module..." -ForegroundColor Cyan
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host "[INFO] ImportExcel module not found. Installing..." -ForegroundColor Yellow
            try {
                Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
                Write-Host "[INFO] ImportExcel module installed successfully." -ForegroundColor Green
            }
            catch {
                Write-Host "❌ Failed to install ImportExcel module: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "   Install manually: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
                return
            }
        } else {
            Write-Host "[INFO] ImportExcel module is available." -ForegroundColor Cyan
        }
        Import-Module ImportExcel -ErrorAction Stop

        # -----------------------------------------------------------------
        # STEP 7: Evaluate criteria against each extension's 1P app config
        # -----------------------------------------------------------------
        Write-Host "[INFO] Evaluating criteria against downloaded configs..." -ForegroundColor Cyan
        Write-Host ""

        # Build the list of column values for the spreadsheet
        # For 'exact' criteria there is one column; for 'contains' there is one column per value
        if ($parsedCriteria.CriteriaType -eq 'contains') {
            $criteriaColumns = @($parsedCriteria.ExpectedValue | ForEach-Object { $_.Trim() })
        } else {
            $criteriaColumns = @($parsedCriteria.ExpectedValue)
        }
        Write-Host "[INFO] Criteria columns for spreadsheet: $($criteriaColumns.Count)" -ForegroundColor Cyan
        foreach ($col in $criteriaColumns) { Write-Host "[INFO]   - '$col'" -ForegroundColor DarkGray }

        $failedExtensions = @()     # Extensions that did NOT satisfy ALL criteria
        $passedExtensions = @()     # Extensions that DID satisfy ALL criteria
        $partialExtensions = @()    # Extensions that satisfied SOME criteria (contains only)
        $errorExtensions  = @()     # Extensions where config could not be downloaded
        $noAppIdExtensions = $extensionsWithoutAppId  # Already collected

        # Rows for the Excel spreadsheet — only extensions with a successfully downloaded config
        $excelRows = @()

        $evalIndex = 0
        foreach ($portalName in ($script:ExtensionCache.Keys | Sort-Object)) {
            if (-not $script:ExtensionCache[$portalName].ContainsKey($Cloud)) { continue }
            $ext = $script:ExtensionCache[$portalName][$Cloud]

            if (-not $ext.OAuthClientId) { continue }

            $evalIndex++
            $appId = $ext.OAuthClientId
            Write-Host "[$evalIndex] Evaluating: $portalName (AppId=$appId)" -ForegroundColor White

            $config = $appConfigs[$appId]
            if ($null -eq $config) {
                Write-Host "    [SKIP] No 1P app config available (download failed or not found)" -ForegroundColor Yellow
                $errorExtensions += @{
                    PortalName = $portalName
                    AppId      = $appId
                    Reason     = 'Config download failed or not found'
                }
                # Excluded from Excel per user preference
                continue
            }

            $evalResult = Test-AppConfigCriteria `
                -AppConfig     $config `
                -CriteriaType  $parsedCriteria.CriteriaType `
                -PropertyPath  $parsedCriteria.PropertyPath `
                -ExpectedValue $parsedCriteria.ExpectedValue

            # Build an Excel row as an ordered hashtable
            $row = [ordered]@{
                'Extension'      = $portalName
                'OAuthClientId'  = $appId
            }

            $foundCount = 0
            $totalCount = $criteriaColumns.Count
            foreach ($col in $criteriaColumns) {
                $valResult = $false
                if ($evalResult.PerValue.Contains($col)) {
                    $valResult = $evalResult.PerValue[$col]
                }
                $row[$col] = if ($valResult) { '✓' } else { '✗' }
                if ($valResult) { $foundCount++ }
            }

            # Add summary column
            $row['Result'] = if ($evalResult.AllSatisfied) { 'PASS' } else { 'FAIL' }
            $row['MatchCount'] = "$foundCount/$totalCount"

            $excelRows += [PSCustomObject]$row

            if ($evalResult.AllSatisfied) {
                Write-Host "    ✓ PASSED ($foundCount/$totalCount)" -ForegroundColor Green
                $passedExtensions += @{ PortalName = $portalName; AppId = $appId }
            } elseif ($foundCount -gt 0) {
                Write-Host "    ~ PARTIAL ($foundCount/$totalCount)" -ForegroundColor Yellow
                $partialExtensions += @{ PortalName = $portalName; AppId = $appId; MatchCount = $foundCount; TotalCount = $totalCount }
            } else {
                Write-Host "    ✗ FAILED (0/$totalCount)" -ForegroundColor Red
                $failedExtensions += @{ PortalName = $portalName; AppId = $appId }
            }
        }

        # -----------------------------------------------------------------
        # STEP 8: Export to Excel
        # -----------------------------------------------------------------
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $excelFileName = "ScanResults_${Cloud}_${timestamp}.xlsx"
        $excelPath = Join-Path -Path (Get-Location) -ChildPath $excelFileName

        Write-Host ""
        Write-Host "[INFO] Exporting results to Excel: $excelPath" -ForegroundColor Cyan

        if ($excelRows.Count -eq 0) {
            Write-Host "⚠ No rows to export (all configs failed to download). Skipping Excel export." -ForegroundColor Yellow
        } else {
            try {
                # Export with formatting
                $excelPkg = $excelRows | Export-Excel -Path $excelPath `
                    -WorksheetName 'ScanResults' `
                    -AutoSize `
                    -AutoFilter `
                    -FreezeTopRow `
                    -BoldTopRow `
                    -TableName 'ScanResults' `
                    -TableStyle Medium6 `
                    -PassThru

                $ws = $excelPkg.Workbook.Worksheets['ScanResults']

                # Color-code the value cells: green for ✓, red for ✗
                $totalRows = $excelRows.Count
                # Columns start at 3 (A=Extension, B=OAuthClientId, then criteria columns)
                $startCol = 3
                $endCol = $startCol + $criteriaColumns.Count - 1

                Write-Host "[INFO] Applying conditional formatting to columns $startCol..$endCol, rows 2..$($totalRows + 1)" -ForegroundColor DarkGray

                for ($r = 2; $r -le ($totalRows + 1); $r++) {
                    for ($c = $startCol; $c -le $endCol; $c++) {
                        $cell = $ws.Cells[$r, $c]
                        $cellValue = $cell.Text
                        if ($cellValue -eq '✓') {
                            $cell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))      # dark green
                            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(198, 239, 206))  # light green bg
                        } elseif ($cellValue -eq '✗') {
                            $cell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(156, 0, 6))      # dark red
                            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 199, 206))  # light red bg
                        }
                    }

                    # Also color the Result column
                    $resultCol = $endCol + 1
                    $resultCell = $ws.Cells[$r, $resultCol]
                    if ($resultCell.Text -eq 'PASS') {
                        $resultCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0, 128, 0))
                        $resultCell.Style.Font.Bold = $true
                    } elseif ($resultCell.Text -eq 'FAIL') {
                        $resultCell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(156, 0, 6))
                        $resultCell.Style.Font.Bold = $true
                    }
                }

                # Center-align the checkmark columns
                for ($c = $startCol; $c -le $endCol; $c++) {
                    $ws.Column($c).Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                }

                $excelPkg.Save()
                $excelPkg.Dispose()

                Write-Host "✅ Excel file saved: $excelPath" -ForegroundColor Green
            }
            catch {
                Write-Host "❌ Failed to export Excel: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "    Stack: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
            }
        }

        # -----------------------------------------------------------------
        # STEP 9: Print console summary report
        # -----------------------------------------------------------------
        Write-Host ""
        Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║              SCAN RESULTS SUMMARY                ║" -ForegroundColor Cyan
        Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Cloud     : $Cloud" -ForegroundColor White
        Write-Host "Criteria  : $Criteria" -ForegroundColor White
        Write-Host "Total ext : $($extensionsInCloud.Count)" -ForegroundColor White
        Write-Host ""

        # Passed
        Write-Host "✓ PASSED ($($passedExtensions.Count) extensions satisfy ALL criteria)" -ForegroundColor Green
        if ($passedExtensions.Count -gt 0 -and $passedExtensions.Count -le 20) {
            foreach ($ext in ($passedExtensions | Sort-Object { $_.PortalName })) {
                Write-Host "    $($ext.PortalName)  (AppId: $($ext.AppId))" -ForegroundColor Green
            }
        } elseif ($passedExtensions.Count -gt 20) {
            Write-Host "    (Too many to list — $($passedExtensions.Count) extensions passed)" -ForegroundColor Green
        }

        # Partial (contains only)
        if ($partialExtensions.Count -gt 0) {
            Write-Host ""
            Write-Host "~ PARTIAL ($($partialExtensions.Count) extensions satisfy SOME but not all criteria)" -ForegroundColor Yellow
            foreach ($ext in ($partialExtensions | Sort-Object { $_.PortalName })) {
                Write-Host "    $($ext.PortalName)  (AppId: $($ext.AppId)) — $($ext.MatchCount)/$($ext.TotalCount) values matched" -ForegroundColor Yellow
            }
        }

        # Failed
        Write-Host ""
        Write-Host "✗ FAILED ($($failedExtensions.Count) extensions do NOT satisfy ANY criteria)" -ForegroundColor Red
        if ($failedExtensions.Count -gt 0) {
            foreach ($ext in ($failedExtensions | Sort-Object { $_.PortalName })) {
                $appConfigUrl = "https://msazure.visualstudio.com/One/_git/AAD-FirstPartyApps?path=/Customers/Configs/AppReg/$($ext.AppId)/AppReg.Parameters.$cloudFileName.json"
                Write-Host "    $($ext.PortalName)  (AppId: $($ext.AppId))" -ForegroundColor Red
                Write-Host "      $appConfigUrl" -ForegroundColor DarkGray
            }
        }

        # Errors
        if ($errorExtensions.Count -gt 0) {
            Write-Host ""
            Write-Host "⚠ ERRORS ($($errorExtensions.Count) extensions could not be evaluated — excluded from Excel)" -ForegroundColor Yellow
            foreach ($ext in ($errorExtensions | Sort-Object { $_.PortalName })) {
                Write-Host "    $($ext.PortalName)  (AppId: $($ext.AppId)) — $($ext.Reason)" -ForegroundColor Yellow
            }
        }

        # No AppId
        if ($noAppIdExtensions.Count -gt 0) {
            Write-Host ""
            Write-Host "ℹ NO APP ID ($($noAppIdExtensions.Count) extensions have no OAuthClientId — skipped)" -ForegroundColor Gray
            foreach ($name in ($noAppIdExtensions | Sort-Object)) {
                Write-Host "    $name" -ForegroundColor Gray
            }
        }

        Write-Host ""
        if ($excelRows.Count -gt 0) {
            Write-Host "📊 Results exported to: $excelPath" -ForegroundColor Cyan
        }
        Write-Host "Scan completed." -ForegroundColor Cyan

        # Return a structured result object for programmatic use
        return [PSCustomObject]@{
            Cloud              = $Cloud
            Criteria           = $Criteria
            TotalExtensions    = $extensionsInCloud.Count
            PassedCount        = $passedExtensions.Count
            PartialCount       = $partialExtensions.Count
            FailedCount        = $failedExtensions.Count
            ErrorCount         = $errorExtensions.Count
            NoAppIdCount       = $noAppIdExtensions.Count
            PassedExtensions   = $passedExtensions
            PartialExtensions  = $partialExtensions
            FailedExtensions   = $failedExtensions
            ErrorExtensions    = $errorExtensions
            NoAppIdExtensions  = $noAppIdExtensions
            ExcelPath          = if ($excelRows.Count -gt 0) { $excelPath } else { $null }
        }
    }
    catch {
        Write-Host "❌ 1P App Scan failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
        throw
    }
}

#==============================================================================
# AZURE DEVOPS PIPELINE DEPENDENCY MAPPING
#==============================================================================
#
# This section contains hardcoded dependency mappings for Azure DevOps pipeline
# stages and jobs. These mappings represent the build and test pipeline structure
# for the Azure Portal project, defining prerequisites and dependencies between
# different build stages and test jobs.
#
# PIPELINE STRUCTURE:
# 1. Build Stages: build_debug, build_retail, build_debug_devoptimized
# 2. Test Stages: Required_Tests, Optional_Tests, Quarantined_Tests  
# 3. Reporting: PublishMergedCodeCoverageReport
# 4. Finalization: Tag_LKG_and_Merge
#
# DEPENDENCY FLOW:
# Build → Tests → Coverage Report → Tag & Merge
#
# Each entry contains:
# - depends_on: Array of prerequisite stages/jobs that must complete first
# - prerequisite_for: Array of stages/jobs that depend on this one
#
# This data is used by pipeline orchestration tools to understand the
# complex dependency relationships in the Azure Portal build system.
#==============================================================================

$deps = @{
  # -------------------------
  # TEMPLATE / EXTERNAL STAGES (present as dependencies in this YAML but defined elsewhere)
  # -------------------------
  'stage:build_debug' = @{
    depends_on        = @()   # not defined here
    prerequisite_for  = @(
      'stage:Required_Tests',
      'stage:Optional_Tests',
      'stage:Quarantined_Tests',
      'stage:Tag_LKG_and_Merge'
    )
  }

  'stage:build_retail' = @{
    depends_on        = @()   # not defined here
    prerequisite_for  = @(
      'stage:Tag_LKG_and_Merge'
      # Also a prerequisite_for Required/Optional/Quarantined *if* buildFlavorForTests=retail,
      # but this YAML expresses that via a parameterized dependsOn (build_${{ parameters.buildFlavorForTests }}),
      # so we keep the concrete debug edges above and note this caveat.
    )
  }

  'stage:build_debug_devoptimized' = @{
    depends_on        = @()   # not referenced by explicit dependsOn in pasted snippet
    prerequisite_for  = @()
  }

  'stage:sdl_sources' = @{
    depends_on        = @()   # template-defined
    prerequisite_for  = @('stage:Tag_LKG_and_Merge')
  }

  # -------------------------
  # STAGES (explicit in YAML)
  # -------------------------
  'stage:Required_Tests' = @{
    depends_on        = @('stage:build_debug')  # parameterized in YAML; assumes buildFlavorForTests=debug
    prerequisite_for  = @(
      'stage:PublishMergedCodeCoverageReport',
      'stage:Tag_LKG_and_Merge'
    )
  }

  'stage:Optional_Tests' = @{
    depends_on        = @('stage:build_debug')  # parameterized; assumes debug
    prerequisite_for  = @('stage:PublishMergedCodeCoverageReport')
  }

  'stage:Quarantined_Tests' = @{
    depends_on        = @('stage:build_debug')  # parameterized; assumes debug
    prerequisite_for  = @('stage:PublishMergedCodeCoverageReport')
  }

  'stage:PublishMergedCodeCoverageReport' = @{
    depends_on        = @(
      'stage:Required_Tests',
      'stage:Optional_Tests',
      'stage:Quarantined_Tests'
    )
    prerequisite_for  = @()
  }

  'stage:Tag_LKG_and_Merge' = @{
    depends_on        = @(
      'stage:Required_Tests',
      'stage:sdl_sources',
      'stage:build_debug',
      'stage:build_retail'
    )
    prerequisite_for  = @()
  }

  # -------------------------
  # JOBS — Required_Tests
  # -------------------------
  'stage:Required_Tests/job:RunThresholdTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunControlsTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunShellTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunOneStbTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunQunitChromeTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunQunitFirefoxTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunMsPortalFxTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunShellTsTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunShellTSPlaywrightTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunDxTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunOneStbTsTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunUnitTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunLoginTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }

  'stage:Required_Tests/job:RunControlsCompat10DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunControlsCompat30DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunControlsCompat120DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }

  'stage:Required_Tests/job:RunShellCompat10DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunShellCompat30DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunShellCompat120DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }

  'stage:Required_Tests/job:RunShellTSCompat10DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunShellTSCompat30DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunShellTSCompat120DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }

  'stage:Required_Tests/job:RunShellTSPlaywrightCompat10DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunShellTSPlaywrightCompat120DaysTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }

  'stage:Required_Tests/job:RunScreenshotTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunSdkv2Tests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:RunSdkv2TemplateTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }

  # React artifact job + dependent jobs
  'stage:Required_Tests/job:React_View_Copy_Files' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @(
      'stage:Required_Tests/job:AzurePortalReactViewUtCloudTests',
      'stage:Required_Tests/job:AzurePortalHubsReactViewUtCloudTests'
    )
  }
  'stage:Required_Tests/job:AzurePortalReactViewUtCloudTests' = @{
    depends_on        = @(
      'stage:Required_Tests',
      'stage:Required_Tests/job:React_View_Copy_Files'
    )
    prerequisite_for  = @()
  }
  'stage:Required_Tests/job:AzurePortalHubsReactViewUtCloudTests' = @{
    depends_on        = @(
      'stage:Required_Tests',
      'stage:Required_Tests/job:React_View_Copy_Files'
    )
    prerequisite_for  = @()
  }

  'stage:Required_Tests/job:RunReactShellTests' = @{
    depends_on        = @('stage:Required_Tests')
    prerequisite_for  = @()
  }

  # -------------------------
  # JOBS — Optional_Tests
  # -------------------------
  'stage:Optional_Tests/job:RunSdkInstallerAuthTests' = @{
    depends_on        = @('stage:Optional_Tests')
    prerequisite_for  = @()
  }
  'stage:Optional_Tests/job:RunSdkInstallerTests' = @{
    depends_on        = @('stage:Optional_Tests')
    prerequisite_for  = @()
  }
  'stage:Optional_Tests/job:RunShellTsTestsWithCodeCoverage' = @{
    depends_on        = @('stage:Optional_Tests')
    prerequisite_for  = @()
  }
  'stage:Optional_Tests/job:RunShellTSReleaseTests' = @{
    depends_on        = @('stage:Optional_Tests')
    prerequisite_for  = @()
  }

  # -------------------------
  # JOBS — Quarantined_Tests
  # -------------------------
  'stage:Quarantined_Tests/job:RunControlsCompat10DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunControlsCompat30DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunControlsCompat120DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellCompat10DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellCompat30DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellCompat120DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellTSCompat10DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellTSCompat30DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellTSCompat120DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellTSPlaywrightCompat10DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellTSPlaywrightCompat120DaysQuarantineTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunScreenshotQuarantinedTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellQuarantinedTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunOneStbQuarantinedTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunLoginQuarantinedTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunQunitQuarantinedTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunQunitFirefoxQuarantinedTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellTSQuarantinedTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunShellTSPlaywrightQuarantinedTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }
  'stage:Quarantined_Tests/job:RunOneStbTSQuarantinedTests' = @{
    depends_on        = @('stage:Quarantined_Tests')
    prerequisite_for  = @()
  }

  # -------------------------
  # JOBS — PublishMergedCodeCoverageReport
  # -------------------------
  'stage:PublishMergedCodeCoverageReport/job:PublishMergedCodeCoverageReport' = @{
    depends_on        = @('stage:PublishMergedCodeCoverageReport')
    prerequisite_for  = @()
  }

  # -------------------------
  # JOBS — Tag_LKG_and_Merge
  # -------------------------
  'stage:Tag_LKG_and_Merge/job:TagAndMerge' = @{
    depends_on        = @('stage:Tag_LKG_and_Merge')
    prerequisite_for  = @()
  }
}

#==============================================================================
# REQUIRED TESTS JOB MAPPING
#==============================================================================
#
# Comprehensive mapping of all test jobs in the Required_Tests stage.
# This mapping provides the bridge between internal job names and 
# user-friendly display names, along with their corresponding skip parameters.
#
# STRUCTURE:
# - Key: Internal job name used in Azure DevOps pipelines
# - displayName: Human-readable name shown in UI
# - skipParameter: Variable name used to skip this specific test
# - sortOrder: Display order in test selection UI (lower = higher priority)
#
# TEST CATEGORIES:
# 1. Core Tests (1-13): Essential functionality tests
# 2. Compatibility Tests (14-24): Backward compatibility verification  
# 3. Specialized Tests (25-31): Screenshot, SDK, React components
#
# This mapping is used by the interactive test selection UI to present
# tests in a logical order and map user selections to pipeline parameters.
#==============================================================================

# Hardcoded map of Required_Tests jobs from the provided YAML
# key   = job name
# value = @{ displayName = '...'; skipParameter = '...' }
$RequiredTestsJobMap = @{
  'RunThresholdTests' = @{
    displayName   = 'Threshold Tests'
    skipParameter = 'SkipThresholdTests'
    sortOrder     = 1
  }
  'RunControlsTests' = @{
    displayName   = 'Controls Tests'
    skipParameter = 'SkipControlsTests'
    sortOrder     = 2
  }
  'RunShellTests' = @{
    displayName   = 'Shell Tests'
    skipParameter = 'SkipShellTests'
    sortOrder     = 3
  }
  'RunOneStbTests' = @{
    displayName   = 'OneStb Tests'
    skipParameter = 'SkipOneStbTests'
    sortOrder     = 4
  }
  'RunQunitChromeTests' = @{
    displayName   = 'QUnit Chrome Tests'
    skipParameter = 'SkipQunitChrome'
    sortOrder     = 5
  }
  'RunQunitFirefoxTests' = @{
    displayName   = 'QUnit Firefox Tests'
    skipParameter = 'SkipQunitFirefoxTests'
    sortOrder     = 6
  }
  'RunMsPortalFxTests' = @{
    displayName   = 'MsPortalFx Tests'
    skipParameter = 'SkipMsPortalFxTests'
    sortOrder     = 7
  }
  'RunShellTsTests' = @{
    displayName   = 'ShellTS Tests'
    skipParameter = 'SkipShellTSTests'
    sortOrder     = 8
  }
  'RunShellTSPlaywrightTests' = @{
    displayName   = 'ShellTS Playwright Tests'
    skipParameter = 'SkipShellTSPlaywrightTests'
    sortOrder     = 9
  }
  'RunDxTests' = @{
    displayName   = 'Dx Tests'
    skipParameter = 'SkipDxTests'
    sortOrder     = 10
  }
  'RunOneStbTsTests' = @{
    displayName   = 'OneStbTS Tests'
    skipParameter = 'SkipOneStbTsTests'
    sortOrder     = 11
  }
  'RunUnitTests' = @{
    displayName   = 'Unit Tests'
    skipParameter = 'SkipUnitTests'
    sortOrder     = 12
  }
  'RunLoginTests' = @{
    displayName   = 'Login Tests'
    skipParameter = 'SkipLoginTests'
    sortOrder     = 13
  }
  'RunControlsCompat10DaysTests' = @{
    displayName   = 'Controls Compat 10d Tests'
    skipParameter = 'SkipControlsCompat10DaysTests'
    sortOrder     = 14
  }
  'RunControlsCompat30DaysTests' = @{
    displayName   = 'Controls Compat 30d Tests'
    skipParameter = 'SkipControlsCompat30DaysTests'
    sortOrder     = 15
  }
  'RunControlsCompat120DaysTests' = @{
    displayName   = 'Controls Compat 120d Tests'
    skipParameter = 'SkipControlsCompat120DaysTests'
    sortOrder     = 16
  }
  'RunShellCompat10DaysTests' = @{
    displayName   = 'Shell Compat 10d Tests'
    skipParameter = 'SkipShellCompat10DaysTests'
    sortOrder     = 17
  }
  'RunShellCompat30DaysTests' = @{
    displayName   = 'Shell Compat 30d Tests'
    skipParameter = 'SkipShellCompat30DaysTests'
    sortOrder     = 18
  }
  'RunShellCompat120DaysTests' = @{
    displayName   = 'Shell Compat 120d Tests'
    skipParameter = 'SkipShellCompat120DaysTests'
    sortOrder     = 19
  }
  'RunShellTSCompat10DaysTests' = @{
    displayName   = 'ShellTS Compat 10d Tests'
    skipParameter = 'SkipShellTSCompat10DaysTests'
    sortOrder     = 20
  }
  'RunShellTSCompat30DaysTests' = @{
    displayName   = 'ShellTS Compat 30d Tests'
    skipParameter = 'SkipShellTSCompat30DaysTests'
    sortOrder     = 21
  }
  'RunShellTSCompat120DaysTests' = @{
    displayName   = 'ShellTS Compat 120d Tests'
    skipParameter = 'SkipShellTSCompat120DaysTests'
    sortOrder     = 22
  }
  'RunShellTSPlaywrightCompat10DaysTests' = @{
    displayName   = 'ShellTS Playwright Compat 10d Tests'
    skipParameter = 'SkipShellTSPlaywrightCompat10DaysTests'
    sortOrder     = 23
  }
  'RunShellTSPlaywrightCompat120DaysTests' = @{
    displayName   = 'ShellTS Playwright Compat 120d Tests'
    skipParameter = 'SkipShellTSPlaywrightCompat120DaysTests'
    sortOrder     = 24
  }
  'RunScreenshotTests' = @{
    displayName   = 'ScreenshotTests'
    skipParameter = 'SkipScreenshotTests'
    sortOrder     = 25
  }
  'RunSdkv2Tests' = @{
    displayName   = 'SDKv2Tests'
    skipParameter = 'SkipSdkv2Tests'
    sortOrder     = 26
  }
  'RunSdkv2TemplateTests' = @{
    displayName   = 'SDKv2TemplateTests'
    skipParameter = 'SkipSdkv2TemplateTests'
    sortOrder     = 27
  }
  'React_View_Copy_Files' = @{
    displayName   = 'Publish Pipeline Artifact'
    skipParameter = $null  # This job doesn't have a single skip parameter mapping
    sortOrder     = 28
  }
  'RunReactShellTests' = @{
    displayName   = 'ReactShellTests'
    skipParameter = 'SkipReactShellTests'
    sortOrder     = 29
  }
  'AzurePortalReactViewUtCloudTests' = @{
    displayName   = 'AzurePortal-ReactView-ut-CloudTest'
    skipParameter = 'SkipReactViewTests'
    sortOrder     = 30
  }
  'AzurePortalHubsReactViewUtCloudTests' = @{
    displayName   = 'AzurePortal-Hubs-ReactView-ut-CloudTest'
    skipParameter = 'SkipHubsReactViewTests'
    sortOrder     = 31
  }
}

#==============================================================================
# PIPELINE STAGE MAPPING
#==============================================================================
#
# Defines the mapping between internal stage names and their display properties.
# Each stage represents a major phase in the Azure Portal build pipeline.
#
# STAGE TYPES:
# - Build Stages: Compile and prepare artifacts for different configurations
# - Test Stages: Execute various test suites (required, optional, quarantined)
# - Report Stage: Aggregate code coverage data from all test stages  
# - Finalization: Tag successful builds and merge to main branches
#
# CONDITIONS:
# Each stage includes condition logic that determines when the stage should run
# based on pipeline variables, previous stage results, and user configuration.
#
#==============================================================================

# Hardcoded map of stages from the provided YAML
# key   = stage name
# value = @{ displayName = '...'; condition = '...' }

$StageMap = @{
  'build_retail' = @{
    displayName = 'Build (Retail)'
    condition   = $null   # condition defined in template, not in this YAML
  }
  'build_debug' = @{
    displayName = 'Build (Debug)'
    condition   = $null   # condition defined in template
  }
  'build_debug_devoptimized' = @{
    displayName = 'Build (Debug DevOptimized)'
    condition   = $null
  }
  'Required_Tests' = @{
    displayName = 'Required Tests'
    condition   = "and(succeeded(), ne(variables['SkipRequiredTests'], 'true'))"
  }
  'Optional_Tests' = @{
    displayName = 'Optional Tests'
    condition   = "and(succeeded(), ne(variables['SkipOptionalTests'], 'true'))"
  }
  'Quarantined_Tests' = @{
    displayName = 'Quarantined Tests'
    condition   = "and(succeeded(), ne(variables['SkipQuarantinedTests'], 'true'))"
  }
  'PublishMergedCodeCoverageReport' = @{
    displayName = 'Publish Merged Code Coverage Report'
    condition   = "or(succeeded(), failed())"
  }
  'Tag_LKG_and_Merge' = @{
    displayName = 'Tag LKG'
    condition   = "or(
      and(
        succeeded(),
        and(
          or(
            eq(variables['Build.SourceBranch'], 'refs/heads/dev'),
            eq(variables['Build.SourceBranch'], 'refs/tags/LKG')
          ),
          ne(variables['SkipTagAndMerge'], 'true')
        )
      ),
      eq(variables['ForceLkgTag'], 'true')
    )"
  }
}

function Invoke-TestRunRequest {
    <#
    .SYNOPSIS
        Hardcoded PowerShell equivalent of submitTestRunRequest(data) in main.js.
    .DESCRIPTION
        Sends a POST to:
          https://dev.azure.com/msazure/{projectGuid}/_apis/pipelines/{pipelineId}/runs?api-version=6.0-preview.1
        Parses response JSON and returns it. Prints a build URL when successful.
    .PARAMETER SkipRequiredTests
        Whether to skip required tests. Defaults to 'false'.
    .PARAMETER SkipQuarantinedTests
        Whether to skip quarantined tests. Defaults to 'true'.
    .PARAMETER SkipOptionalTests
        Whether to skip optional tests. Defaults to 'true'.
    .PARAMETER BranchName
        The branch name to run tests against. Defaults to 'dev'.
    .EXAMPLE
        Invoke-TestRunRequest -SkipRequiredTests 'true' -SkipLoginTests 'false'
    .EXAMPLE
        Invoke-TestRunRequest -BranchName 'main' -SkipRequiredTests 'false'
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string] $SkipRequiredTests = 'false',
        
        [Parameter()]
        [string] $SkipQuarantinedTests = 'true',
        
        [Parameter()]
        [string] $SkipOptionalTests = 'true',
        
        # Individual test control parameters - these provide granular control
        # over specific test suites within the Required_Tests stage
        [Parameter()]
        [string] $SkipControlsCompat10DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipControlsCompat30DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipControlsCompat120DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipControlsTests = 'true',
        
        [Parameter()]
        [string] $SkipDxTests = 'true',
        
        [Parameter()]
        [string] $SkipHubsReactViewFluentv8Tests = 'true',
        
        [Parameter()]
        [string] $SkipHubsReactViewTests = 'true',
        
        [Parameter()]
        [string] $SkipLoginTests = 'true',
        
        [Parameter()]
        [string] $SkipMsPortalFxTests = 'true',
        
        [Parameter()]
        [string] $SkipOneStbTests = 'true',
        
        [Parameter()]
        [string] $SkipOneStbTsTests = 'true',
        
        [Parameter()]
        [string] $SkipQunitChrome = 'true',
        
        [Parameter()]
        [string] $SkipQunitFirefoxTests = 'true',
        
        [Parameter()]
        [string] $SkipQunitTests = 'true',
        
        [Parameter()]
        [string] $SkipReactShellTests = 'true',
        
        [Parameter()]
        [string] $SkipReactViewFluentv8Tests = 'true',
        
        [Parameter()]
        [string] $SkipReactViewTests = 'true',
        
        [Parameter()]
        [string] $SkipScreenshotTests = 'true',
        
        [Parameter()]
        [string] $SkipShellCompat10DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipShellCompat30DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipShellCompat120DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipShellTests = 'true',
        
        [Parameter()]
        [string] $SkipShellTSCompat10DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipShellTSCompat30DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipShellTSCompat120DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipShellTSPlaywrightCompat10DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipShellTSPlaywrightCompat120DaysTests = 'true',
        
        [Parameter()]
        [string] $SkipShellTSPlaywrightTests = 'true',
        
        [Parameter()]
        [string] $SkipShellTSTests = 'true',
        
        [Parameter()]
        [string] $SkipSdkv2TemplateTests = 'true',
        
        [Parameter()]
        [string] $SkipSdkv2Tests = 'true',
        
        [Parameter()]
        [string] $SkipThresholdTests = 'true',
        
        [Parameter()]
        [string] $SkipUnitTests = 'true',
        
        [Parameter()]
        [string] $RunControlsCompat30DaysTests = 'true',
        
        [Parameter()]
        [string] $BranchName = 'dev'
    )

    # --- Hardcoded constants from main.js ---
    $Organization = 'https://dev.azure.com/msazure'
    $ProjectGuid  = 'b32aa71e-8ed2-41b2-9d77-5bc261222004'  # AdoAzurePortalProjectGuid :contentReference[oaicite:2]{index=2}
    $PipelineId   = '198960'                                 # AdoAzureUXPortalFxOnDemandDevCiPipelineId :contentReference[oaicite:3]{index=3}
    $ApiVersion   = '7.1'                                    # Simplified API version without preview suffix

    # Build the ref name from the branch parameter
    $RefName = "refs/heads/$BranchName"                      # matches JS: refs/heads/${branchName} :contentReference[oaicite:5]{index=5}

    # --- Hardcoded request body (keep it literal for the test) ---
    # JS sends JSON.stringify(data) with stagesToSkip/resources/variables/templateParameters :contentReference[oaicite:6]{index=6}
    $BodyObject = [ordered]@{
        stagesToSkip = @(
            # Keep empty for a basic test. Later you can populate with stage names like "Optional_Tests", etc.
        )
        resources = @{
            repositories = @{
                self = @{
                    refName = $RefName
                }
            }
        }
        variables = @{
            CloudTestAccount = @{ value = 'azureportal' }
            RetryOnFailureMode = @{ value = 'None' }

            # Test control variables
            RunControlsCompat30DaysTests = @{ value = $RunControlsCompat30DaysTests }

            # Core pipeline control - determines which test stages run
            SkipRequiredTests     = @{ value = $SkipRequiredTests }
            SkipQuarantinedTests  = @{ value = $SkipQuarantinedTests }
            SkipOptionalTests     = @{ value = $SkipOptionalTests }

            # Individual test skip parameters (granular control within Required_Tests stage)
            # NOTE: Some parameters are temporarily commented out as they are not settable at queue time
            # SkipControlsCompat10DaysTests  = @{ value = $SkipControlsCompat10DaysTests }  # TODO: Temporarily commented - not settable at queue time
            # SkipControlsCompat30DaysTests  = @{ value = $SkipControlsCompat30DaysTests }  # TODO: Temporarily commented - not settable at queue time
            # SkipControlsCompat120DaysTests = @{ value = $SkipControlsCompat120DaysTests } # TODO: Temporarily commented - not settable at queue time
            SkipControlsTests     = @{ value = $SkipControlsTests }
            # SkipDxTests         = @{ value = $SkipDxTests }                               # TODO: Temporarily commented - not settable at queue time
            # SkipHubsReactViewFluentv8Tests = @{ value = $SkipHubsReactViewFluentv8Tests } # TODO: Temporarily commented - not settable at queue time
            # SkipHubsReactViewTests       = @{ value = $SkipHubsReactViewTests }           # TODO: Temporarily commented - not settable at queue time
            SkipLoginTests        = @{ value = $SkipLoginTests }
            SkipMsPortalFxTests   = @{ value = $SkipMsPortalFxTests }
            SkipOneStbTests       = @{ value = $SkipOneStbTests }
            SkipOneStbTsTests     = @{ value = $SkipOneStbTsTests }
            SkipQunitChrome       = @{ value = $SkipQunitChrome }
            SkipQunitFirefoxTests = @{ value = $SkipQunitFirefoxTests }
            SkipQunitTests        = @{ value = $SkipQunitTests }
            # SkipReactShellTests    = @{ value = $SkipReactShellTests }                    # TODO: Temporarily commented - not settable at queue time
            # SkipReactViewFluentv8Tests = @{ value = $SkipReactViewFluentv8Tests }         # TODO: Temporarily commented - not settable at queue time
            # SkipReactViewTests    = @{ value = $SkipReactViewTests }                      # TODO: Temporarily commented - not settable at queue time
            # SkipScreenshotTests    = @{ value = $SkipScreenshotTests }                   # TODO: Temporarily commented - not settable at queue time
            # SkipShellCompat10DaysTests = @{ value = $SkipShellCompat10DaysTests }        # TODO: Temporarily commented - not settable at queue time
            SkipShellCompat30DaysTests = @{ value = $SkipShellCompat30DaysTests }
            # SkipShellCompat120DaysTests = @{ value = $SkipShellCompat120DaysTests }      # TODO: Temporarily commented - not settable at queue time
            SkipShellTests            = @{ value = $SkipShellTests }
            # SkipShellTSCompat10DaysTests = @{ value = $SkipShellTSCompat10DaysTests }    # TODO: Temporarily commented - not settable at queue time
            SkipShellTSCompat30DaysTests = @{ value = $SkipShellTSCompat30DaysTests }
            # SkipShellTSCompat120DaysTests = @{ value = $SkipShellTSCompat120DaysTests }  # TODO: Temporarily commented - not settable at queue time
            # SkipShellTSPlaywrightCompat10DaysTests = @{ value = $SkipShellTSPlaywrightCompat10DaysTests } # TODO: Temporarily commented - not settable at queue time
            # SkipShellTSPlaywrightCompat120DaysTests = @{ value = $SkipShellTSPlaywrightCompat120DaysTests } # TODO: Temporarily commented - not settable at queue time
            # SkipShellTSPlaywrightTests = @{ value = $SkipShellTSPlaywrightTests }        # TODO: Temporarily commented - not settable at queue time
            SkipShellTSTests          = @{ value = $SkipShellTSTests }
            # SkipSdkv2TemplateTests = @{ value = $SkipSdkv2TemplateTests }                # TODO: Temporarily commented - not settable at queue time
            # SkipSdkv2Tests          = @{ value = $SkipSdkv2Tests }                       # TODO: Temporarily commented - not settable at queue time
            SkipThresholdTests        = @{ value = $SkipThresholdTests }
            SkipUnitTests             = @{ value = $SkipUnitTests }
        }
        templateParameters = @{
            debug = 'False'
            buildFlavorForTests = 'debug'
        }
    }

    $json = $BodyObject | ConvertTo-Json -Depth 20

    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        $json | Out-File -FilePath $tempFile -Encoding utf8

        # Use helper function to handle authentication automatically
        $respJson = Invoke-WithAzDevOpsAuth -ScriptBlock {
            # Equivalent POST to the JS URL: /_apis/pipelines/{pipelineId}/runs?api-version=6.0-preview.1 :contentReference[oaicite:8]{index=8}
            az devops invoke `
                --organization $Organization `
                --area pipelines `
                --resource runs `
                --route-parameters "project=$ProjectGuid" "pipelineId=$PipelineId" `
                --http-method POST `
                --api-version $ApiVersion `
                --in-file $tempFile `
                --only-show-errors `
                -o json
        }

        if ([string]::IsNullOrWhiteSpace($respJson)) {
            # Pipeline submission failed - this usually indicates authentication,
            # permission, or configuration issues (invalid project GUID, pipeline ID, etc.)
            throw "Pipeline submission failed. No response received."
        }

        $resp = $respJson | ConvertFrom-Json

        # JS expects response.id and then navigates to build results URL :contentReference[oaicite:9]{index=9}
        if ($null -ne $resp.id) {
            $buildUrl = "https://dev.azure.com/msazure/$ProjectGuid/_build/results?buildId=$($resp.id)"
            Write-Host $buildUrl -ForegroundColor Green
        } else {
            Write-Host "Request returned JSON but no .id was found." -ForegroundColor Yellow
            $resp | ConvertTo-Json -Depth 10 | Write-Host
        }

        return $resp
    }
    finally {
        Remove-Item $tempFile -ErrorAction SilentlyContinue
    }
}

function Run-OnDemandCiTestWithUi {
    <#
    .SYNOPSIS
        Shows a simple checkbox UI to select which tests to run.
    .OUTPUTS
        PSCustomObject with BranchName, Mode, SelectedTests (string[])
        Returns $null if user cancels/closes.
        SelectedTests will contain the underlying test keys (e.g., 'RunLoginTests') rather than display names.
    #>

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Create test items with both display names and underlying values
    $testItems = @()
    foreach ($key in ($RequiredTestsJobMap.Keys | Sort-Object)) {
        $displayName = $RequiredTestsJobMap[$key].displayName
        $skipParameter = $RequiredTestsJobMap[$key].skipParameter
        $sortOrder = if ($RequiredTestsJobMap[$key].sortOrder) { $RequiredTestsJobMap[$key].sortOrder } else { 999 }
        $testItems += [PSCustomObject]@{
            DisplayName = $displayName
            Value = $key
            SkipParameter = $skipParameter
            SortOrder = $sortOrder
        }
    }
    
    # Sort by custom sort order, then by display name for items without sort order
    $testItems = $testItems | Sort-Object SortOrder

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Tests to Run"
    $form.Size = New-Object System.Drawing.Size(360, 520)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true

    # Branch label + textbox
    $lblBranch = New-Object System.Windows.Forms.Label
    $lblBranch.Text = "Choose your branch:"
    $lblBranch.Location = New-Object System.Drawing.Point(12, 12)
    $lblBranch.AutoSize = $true
    $form.Controls.Add($lblBranch)

    $txtBranch = New-Object System.Windows.Forms.TextBox
    $txtBranch.Location = New-Object System.Drawing.Point(12, 34)
    $txtBranch.Size = New-Object System.Drawing.Size(320, 20)
    $txtBranch.Text = (git branch --show-current)
    $form.Controls.Add($txtBranch)

    # Mode radio buttons
    $grpMode = New-Object System.Windows.Forms.GroupBox
    $grpMode.Text = "Mode"
    $grpMode.Location = New-Object System.Drawing.Point(12, 65)
    $grpMode.Size = New-Object System.Drawing.Size(320, 55)
    $form.Controls.Add($grpMode)

    $rbBasic = New-Object System.Windows.Forms.RadioButton
    $rbBasic.Text = "Basic"
    $rbBasic.Location = New-Object System.Drawing.Point(12, 22)
    $rbBasic.Checked = $true
    $grpMode.Controls.Add($rbBasic)

    $rbAdvanced = New-Object System.Windows.Forms.RadioButton
    $rbAdvanced.Text = "Advanced"
    $rbAdvanced.Location = New-Object System.Drawing.Point(120, 22)
    $grpMode.Controls.Add($rbAdvanced)

    # Checklist box
    $grpTests = New-Object System.Windows.Forms.GroupBox
    $grpTests.Text = "Tests"
    $grpTests.Location = New-Object System.Drawing.Point(12, 130)
    $grpTests.Size = New-Object System.Drawing.Size(320, 300)
    $form.Controls.Add($grpTests)

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location = New-Object System.Drawing.Point(12, 22)
    $clb.Size = New-Object System.Drawing.Size(295, 260)
    $clb.CheckOnClick = $true
    $clb.DisplayMember = "DisplayName"  # Show the friendly display name
    
    # Add the test items to the checklist
    foreach ($item in $testItems) {
        [void]$clb.Items.Add($item)
    }
    
    $grpTests.Controls.Add($clb)

    # Buttons
    $btnSubmit = New-Object System.Windows.Forms.Button
    $btnSubmit.Text = "Submit"
    $btnSubmit.Location = New-Object System.Drawing.Point(172, 440)
    $btnSubmit.Size = New-Object System.Drawing.Size(75, 28)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(257, 440)
    $btnCancel.Size = New-Object System.Drawing.Size(75, 28)

    $form.Controls.Add($btnSubmit)
    $form.Controls.Add($btnCancel)

    $result = $null

    $btnSubmit.Add_Click({
        $branch = $txtBranch.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($branch)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a branch name.", "Missing branch")
            return
        }

        $mode = if ($rbAdvanced.Checked) { "Advanced" } else { "Basic" }

        # Get the underlying values instead of display names
        $selected = @()
        foreach ($item in $clb.CheckedItems) { 
            $selected += $item.Value  # Use the underlying test key
        }
        
        # Also log to console for debugging
        if ($selected.Count -gt 0) {
            # Convert selected test keys to display names for better messaging
            $displayNames = @()
            foreach ($testKey in $selected) {
                if ($RequiredTestsJobMap.ContainsKey($testKey)) {
                    $displayNames += $RequiredTestsJobMap[$testKey].displayName
                } else {
                    $displayNames += $testKey  # Fallback to key if no display name found
                }
            }
            Write-Host "Running tests: $($displayNames -join ', ') - submitting pipeline request..." -ForegroundColor Cyan
        } else {
            Write-Host "No tests selected" -ForegroundColor Yellow
            return
        }

        $result = [pscustomobject]@{
            BranchName    = $branch
            Mode          = $mode
            SelectedTests = $selected
        }

        # Close the form immediately after validation
        $form.Close()

        # Call Invoke-TestRunRequest with the selected parameters
        try {
            
            # Start with all tests skipped by default
            $testParams = @{
                BranchName = $branch
                SkipRequiredTests = 'false'  # Always run required tests stage
                SkipQuarantinedTests = 'true'
                SkipOptionalTests = 'true'
                SkipCompatTests = 'true'
                
                # Individual test skip parameters - default to true (skip)
                SkipControlsCompat10DaysTests = 'true'
                SkipControlsCompat30DaysTests = 'true'
                SkipControlsCompat120DaysTests = 'true'
                SkipControlsTests = 'true'
                SkipDxTests = 'true'
                SkipHubsReactViewFluentv8Tests = 'true'
                SkipHubsReactViewTests = 'true'
                SkipLoginTests = 'true'
                SkipMsPortalFxTests = 'true'
                SkipOneStbTests = 'true'
                SkipOneStbTsTests = 'true'
                SkipQunitChrome = 'true'
                SkipQunitFirefoxTests = 'true'
                SkipQunitTests = 'true'
                SkipReactShellTests = 'true'
                SkipReactViewFluentv8Tests = 'true'
                SkipReactViewTests = 'true'
                SkipScreenshotTests = 'true'
                SkipShellCompat10DaysTests = 'true'
                SkipShellCompat30DaysTests = 'true'
                SkipShellCompat120DaysTests = 'true'
                SkipShellTests = 'true'
                SkipShellTSCompat10DaysTests = 'true'
                SkipShellTSCompat30DaysTests = 'true'
                SkipShellTSCompat120DaysTests = 'true'
                SkipShellTSPlaywrightCompat10DaysTests = 'true'
                SkipShellTSPlaywrightCompat120DaysTests = 'true'
                SkipShellTSPlaywrightTests = 'true'
                SkipShellTSTests = 'true'
                SkipSdkv2TemplateTests = 'true'
                SkipSdkv2Tests = 'true'
                SkipThresholdTests = 'true'
                SkipUnitTests = 'true'
                RunControlsCompat30DaysTests = 'true'
            }

            # For selected tests, set their skip parameters to 'false' (don't skip them)
            foreach ($testName in $selected) {
                if ($RequiredTestsJobMap.ContainsKey($testName) -and $RequiredTestsJobMap[$testName].skipParameter) {
                    $skipParam = $RequiredTestsJobMap[$testName].skipParameter
                    $testParams[$skipParam] = 'false'
                }
                
                # Special handling for QUnit tests - if any QUnit test is selected, don't skip QUnit entirely
                if ($testName -eq 'RunQunitChromeTests' -or $testName -eq 'RunQunitFirefoxTests') {
                    $testParams['SkipQunitTests'] = 'false'
                }
                
                # Special handling for compat tests - if any compat test is selected, don't skip compat entirely
                if ($testName -match 'Compat.*Tests$') {
                    # Note: SkipCompatTests parameter was removed as it's not needed with granular test control
                }
            }

            # Call the function with splatting
            Invoke-TestRunRequest @testParams
        }
        catch {
            $errorMsg = "Failed to submit test run. Error: $($_.Exception.Message)"
            Write-Host $errorMsg -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $btnCancel.Add_Click({
        $result = $null
        $form.Close()
    })

    # If user closes window via X
    $form.Add_FormClosing({
        if (-not $result) { $result = $null }
    })

    [void]$form.ShowDialog()
    return $result
}

#==============================================================================
# PUBLIC WRAPPER FUNCTIONS
#==============================================================================
#
# This section contains user-facing commands that provide streamlined workflows
# for common Azure DevOps operations. These functions are designed to be:
# - Easy to remember and type (short names, intuitive parameters)
# - Comprehensive in functionality (handle full workflows end-to-end)  
# - Robust with error handling and user feedback
# - Consistent in behavior and output formatting
#
# MAIN CATEGORIES:
# 1. Git Workflow: newbranch, createpr, push
# 2. Work Item Management: createworkitem, bug, pbi, listworkitems
# 3. Pipeline Testing: runtests (via Run-OnDemandCiTestWithUi)
# 4. Extension Lookup: getextensiondetails
#
# All functions include comprehensive help documentation and examples.
# Many support both interactive prompts and parameter-based usage.
#==============================================================================
function newbranch {
    param([Parameter(Mandatory)]$branchName)
    
    if ([string]::IsNullOrWhiteSpace($branchName)) {
        Write-Error "Branch name cannot be empty or null" -ErrorAction Stop
        return
    }
    
    git checkout -b "aksingal/$branchName" origin/dev
    git pull
    git rebase

    # run clean.ps1 from the repo
    $repoRoot = git rev-parse --show-toplevel
    if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($repoRoot)) {
        $cleanScriptPath = Join-Path $repoRoot 'tools\PathScripts\clean.ps1'
        if (Test-Path $cleanScriptPath) {
            & $cleanScriptPath
        } else {
            Write-Warning "Clean script not found at path: $cleanScriptPath"
        }
    } else {
        Write-Warning "Unable to determine repository root. Skipping clean.ps1 invocation."
    }

    # stop IIS
    net stop w3svc

    # kill common processes that might interfere with file locks or builds
    $imageNames = @(
        'w3wp.exe',
        'MSBuild.exe',
        'IISExpress.exe',
        'ChromeDriver.exe',
        'MicrosoftWebDriver.exe',
        'iojs.exe',
        'node.exe',
        'dotnet.exe'
    )

    foreach ($imageName in $imageNames) {
        & taskkill.exe /F /IM $imageName 2>$null | Out-Null
    }

    # run restore_repo.ps1 from the repo
    $restoreScriptPath = Join-Path $repoRoot 'restore_repo.ps1'
    if (Test-Path $restoreScriptPath) {
        & $restoreScriptPath
    } else {
        Write-Warning "Restore script not found at path: $restoreScriptPath"
    }

    # build onestb project
    $buildOneStbPath = Join-Path $repoRoot 'src\Shared\ProjectShortcuts\buildonestb'
    if (Test-Path $buildOneStbPath) {
        dotnet build $buildOneStbPath
    } else {
        Write-Warning "buildonestb path not found at: $buildOneStbPath"
    }

    # start IIS
    net start w3svc

    # Load onestb in the browser to warm up the local dev environment
    Start-Process "https://onestb.cloudapp.net/"
}

function createworkitem {
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

function createpr {
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

function runtests {
    <#
    .SYNOPSIS
        Alias for Run-OnDemandCiTestWithUi - shows test picker dialog.
    .DESCRIPTION
        Shorthand command to open the test selection UI for running Azure DevOps pipeline tests.
    .EXAMPLE
        runtests
    #>
    [CmdletBinding()]
    param()
    
    Run-OnDemandCiTestWithUi
}

function push {
    <#
    .SYNOPSIS
        Amends the last commit with staged changes and force pushes to the remote branch.
    .DESCRIPTION
        This function provides a streamlined workflow for updating the most recent commit with new changes.
        It performs the following operations:
        1. Checks for staged files in the git index
        2. Amends the last commit with staged changes without editing the commit message
        3. Force pushes the amended commit to the remote branch
        
        This is particularly useful during iterative development where you want to refine the same
        logical change without creating multiple commits. The function automatically detects the
        current branch and pushes to the corresponding remote branch.
        
        WARNING: This function uses force push (-f) which rewrites git history. Use with caution
        on shared branches as it may cause issues for other developers who have already pulled
        the original commit.
    .EXAMPLE
        push
        
        Amends the last commit with all staged changes and force pushes to the current branch.
    .EXAMPLE
        # Stage some changes first
        git add file1.txt file2.txt
        push
        
        Adds the specified files to the staging area, then amends the last commit and pushes.
    .NOTES
        - Requires staged files to be present; exits early if no staged changes are found
        - Uses 'git commit --amend --no-edit' to preserve the original commit message
        - Uses 'git push -f origin $currentBranch' for force pushing
        - Provides colored output for success/failure status
        - Automatically determines the current branch name
    #>
    $currentBranch = git branch --show-current
    
    # Check if there are staged files
    $stagedFiles = git diff --cached --name-only
    if (-not $stagedFiles) {
        Write-Host "No staged files found. Nothing to push." -ForegroundColor Yellow
        return
    }
    
    # Prompt for commit message
    $commitMessage = Read-Host "Commit message (press Enter to amend last commit)"
    
    # Create commit or amend based on user input
    if ([string]::IsNullOrWhiteSpace($commitMessage)) {
        # Amend the last commit with staged changes
        git commit --amend --no-edit
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Failed to amend commit" -ForegroundColor Red
            return
        }
        Write-Host "✓ Amended last commit" -ForegroundColor Gray
    } else {
        # Create new commit with provided message
        git commit -m $commitMessage
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Failed to create commit" -ForegroundColor Red
            return
        }
        Write-Host "✓ Created new commit: $commitMessage" -ForegroundColor Gray
    }
    
    # Force push to remote
    git push -f origin $currentBranch
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to push to remote" -ForegroundColor Red
        return
    }
    
    Write-Host "✓ Pushed to $currentBranch" -ForegroundColor Green

    $myPullRequestsUrl = 'https://msazure.visualstudio.com/One/_git/AzureUX-PortalFx/pullrequests?_a=mine'
    Start-Process $myPullRequestsUrl
}

# Query recent work items assigned to me with pagination
function listworkitems {
    <#
    .SYNOPSIS
        Queries and displays recent work items assigned to the current user with pagination.
    .DESCRIPTION
        Retrieves work items assigned to the current user, sorted by creation date (most recent first).
        Displays results in groups of 5 items with interactive navigation.
        
        Navigation:
        - Press 'n' or 'N' to view the next 5 items
        - Press any other key to exit
        
        Each work item displays: ID, Title, Type, State, and Creation Date.
    .EXAMPLE
        listworkitems
        
        Displays the first 5 work items assigned to you, with option to navigate through more.
    .EXAMPLE
        listworkitems -Verbose
        
        Displays work items with detailed timing information.
    .NOTES
        - Uses Azure DevOps CLI (az devops) to query work items
        - Automatically handles authentication via Ensure-AzDevOpsLogin
        - Results are sorted by System.CreatedDate in descending order
        - Shows work items from the 'One' project in the msazure organization
    #>
    [CmdletBinding()]
    param()
    
    $pageSize = 5
    $skip = 0
    $hasMore = $true
    
    Write-Host "`n=== My Recent Work Items ===" -ForegroundColor Cyan
    Write-Host ""
    
    while ($hasMore) {
        try {
            # Build WIQL query to get work items assigned to current user
            $wiql = @"
SELECT [System.Id], [System.Title], [System.WorkItemType], [System.State], [System.CreatedDate]
FROM WorkItems
WHERE [System.AssignedTo] = @Me
ORDER BY [System.CreatedDate] DESC
"@
            
            # Create JSON payload for WIQL query
            $queryPayload = @{
                query = $wiql
            } | ConvertTo-Json -Depth 10
            
            # Create temp file for WIQL query
            $tempFile = [System.IO.Path]::GetTempFileName()
            $queryPayload | Out-File -FilePath $tempFile -Encoding utf8
            
            try {
                if ($VerbosePreference -eq 'Continue') { Write-Host "⏱ Starting WIQL query..." -ForegroundColor Gray }
                $queryStart = Get-Date
                
                # Use helper function to handle authentication automatically
                $queryResult = Invoke-WithAzDevOpsAuth -ScriptBlock {
                    # Execute WIQL query with top 50 limit
                    az devops invoke `
                        --organization 'https://dev.azure.com/msazure' `
                        --area wit `
                        --resource wiql `
                        --route-parameters "project=One" `
                        --query-parameters '$top=50' `
                        --http-method POST `
                        --api-version '7.1' `
                        --media-type 'application/json' `
                        --in-file $tempFile `
                        --only-show-errors | ConvertFrom-Json
                }
                
                $queryEnd = Get-Date
                $queryDuration = ($queryEnd - $queryStart).TotalSeconds
                if ($VerbosePreference -eq 'Continue') { Write-Host "✓ WIQL query completed in $([math]::Round($queryDuration, 2)) seconds" -ForegroundColor Gray }
                
                if (-not $queryResult -or -not $queryResult.workItems) {
                    Write-Host "No work items found assigned to you." -ForegroundColor Yellow
                    break
                }
                
                # Get the work item IDs for this page (API already limited to top 50)
                $totalItems = $queryResult.workItems.Count
                $pageItems = $queryResult.workItems | Select-Object -Skip $skip -First $pageSize
                
                if (-not $pageItems -or $pageItems.Count -eq 0) {
                    Write-Host "No more work items to display." -ForegroundColor Yellow
                    break
                }
                
                # Get detailed work item information using batch query
                $workItemIds = $pageItems | ForEach-Object { $_.id }
                $idsString = $workItemIds -join ','
                
                if ($VerbosePreference -eq 'Continue') { Write-Host "⏱ Fetching details for $($workItemIds.Count) work items in batch..." -ForegroundColor Gray }
                $detailsStart = Get-Date
                
                try {
                    # Optimization: run individual queries in parallel using background jobs
                    $jobs = @()
                    foreach ($id in $workItemIds) {
                        $jobs += Start-Job -ScriptBlock {
                            param($id, $org)
                            
                            try {
                                $result = az boards work-item show `
                                    --organization $org `
                                    --id $id `
                                    --fields "System.Id,System.Title,System.State,System.CreatedDate" `
                                    --expand none `
                                    --only-show-errors `
                                    -o json | ConvertFrom-Json

                                return $result
                            }
                            catch {
                                # Auth may be expired, let the parent process handle re-auth
                                return $null
                            }
                        } -ArgumentList $id, 'https://dev.azure.com/msazure'
                    }
                    
                    # Wait for all jobs to complete and collect results
                    $workItems = @()
                    $authRetryNeeded = $false
                    
                    foreach ($job in $jobs) {
                        $result = Receive-Job -Job $job -Wait
                        if ($result) {
                            $workItems += $result
                        } elseif ($result -eq $null -and -not $authRetryNeeded) {
                            $authRetryNeeded = $true
                        }
                        Remove-Job -Job $job
                    }
                    
                    # If some jobs failed due to auth, retry them after re-authenticating
                    if ($authRetryNeeded -and $workItems.Count -lt $workItemIds.Count) {
                        Write-Host "Some work items failed to load, retrying with fresh authentication..." -ForegroundColor Yellow
                        Ensure-AzDevOpsLogin
                        
                        # Retry failed items
                        $missingIds = $workItemIds | Where-Object { $_ -notin ($workItems | ForEach-Object { $_.fields.'System.Id' }) }
                        if ($missingIds) {
                            $retryJobs = @()
                            foreach ($id in $missingIds) {
                                $retryJobs += Start-Job -ScriptBlock {
                                    param($id, $org)
                                    try {
                                        az boards work-item show --organization $org --id $id --fields "System.Id,System.Title,System.State,System.CreatedDate" --expand none --only-show-errors -o json | ConvertFrom-Json
                                    }
                                    catch {
                                        $null
                                    }
                                } -ArgumentList $id, 'https://dev.azure.com/msazure'
                            }
                            
                            foreach ($job in $retryJobs) {
                                $result = Receive-Job -Job $job -Wait
                                if ($result) {
                                    $workItems += $result
                                }
                                Remove-Job -Job $job
                            }
                        }
                    }
                }
                catch {
                    Write-Host "Error fetching work item details: $($_.Exception.Message)" -ForegroundColor Red
                    break
                }
                
                $detailsEnd = Get-Date
                $detailsDuration = ($detailsEnd - $detailsStart).TotalSeconds
                if ($VerbosePreference -eq 'Continue') { Write-Host "✓ All work item details fetched in $([math]::Round($detailsDuration, 2)) seconds" -ForegroundColor Gray }
                
                # Display work items
                $currentPage = [Math]::Floor($skip / $pageSize) + 1
                $startItem = $skip + 1
                $endItem = [Math]::Min($skip + $pageSize, $totalItems)
                
                Write-Host "Page $currentPage - Showing $startItem-$endItem of $totalItems work items:" -ForegroundColor White
                Write-Host ("=" * 80) -ForegroundColor DarkGray
                
                foreach ($item in $workItems) {
                    if (-not $item -or -not $item.fields) {
                        continue
                    }
                    
                    $fields = $item.fields
                    $createdDate = [DateTime]::Parse($fields.'System.CreatedDate').ToString("MM/dd/yyyy")
                    
                    # First line: #ID: Title
                    Write-Host "#" -NoNewline -ForegroundColor Gray
                    Write-Host $fields.'System.Id' -NoNewline -ForegroundColor White
                    Write-Host ": " -NoNewline -ForegroundColor Gray
                    Write-Host $fields.'System.Title' -ForegroundColor White
                    
                    # Second line: State | Created
                    Write-Host "State: " -NoNewline -ForegroundColor Gray
                    
                    # Color code the state based on value
                    $stateColor = switch ($fields.'System.State') {
                        'In Review' { 'DarkYellow' }
                        'Resolved'  { 'Green' }
                        'Done'  { 'Green' }
                        'Removed'  { 'Gray' }
                        default     { 'Red' }
                    }
                    
                    Write-Host $fields.'System.State' -NoNewline -ForegroundColor $stateColor
                    Write-Host " | Created: " -NoNewline -ForegroundColor Gray
                    Write-Host $createdDate -ForegroundColor Yellow
                    
                    # Third line: URL
                    Write-Host "URL: " -NoNewline -ForegroundColor Gray
                    Write-Host "https://msazure.visualstudio.com/One/_workitems/edit/$($fields.'System.Id')" -ForegroundColor Blue
                    
                    Write-Host ""
                }
                
                # Check if there are more items
                $hasMore = ($skip + $pageSize) -lt $totalItems
                
                if ($hasMore) {
                    Write-Host ("=" * 80) -ForegroundColor DarkGray
                    Write-Host "Press 'n' for next page, any other key to exit: " -NoNewline -ForegroundColor Yellow
                    $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").Character.ToString().ToLower()
                    Write-Host $key -ForegroundColor White
                    
                    if ($key -eq 'n') {
                        $skip += $pageSize
                        Write-Host ""
                    } else {
                        Write-Host "Exiting..." -ForegroundColor Gray
                        $hasMore = $false
                    }
                } else {
                    Write-Host ("=" * 80) -ForegroundColor DarkGray
                    Write-Host "End of results. Press any key to exit: " -NoNewline -ForegroundColor Yellow
                    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
                    Write-Host ""
                    $hasMore = $false
                }
            }
            finally {
                Remove-Item $tempFile -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Host "Error retrieving work items: $($_.Exception.Message)" -ForegroundColor Red
            break
        }
    }
}

#==============================================================================
# CONVENIENT ALIASES AND SHORTCUTS
#==============================================================================
#
# Short, memorable aliases for frequently used commands. These provide
# quick access to common operations without sacrificing functionality.
#
# ALIAS CATEGORIES:
# - Git Operations: s (status), nb (new branch), pr (pull request)
# - Work Items: lwi (list work items)  
# - Extensions: getextensiondetails (extension lookup)
#
# Each alias maintains full parameter support and help documentation
# from the original function while providing faster typing.
#==============================================================================

# PR creation alias
function pr {
    createpr @args
}

# New branch creation alias
function nb {
    newbranch @args
}

# Quick git status alias with last commit
function s { 
    git status $args
    Write-Host ""
    Write-Host "Last commit:" -ForegroundColor Yellow
    git log -1 --oneline
}

# list work item alias
function lwi {
    listworkitems @args
}

# extension details alias
function getextensiondetails {
    Get-ExtensionDetails @args
}

# 1P app scan alias
function scanapps {
    Invoke-FirstPartyAppScan @args
}

#==============================================================================
# AUTOMATIC INITIALIZATION
#==============================================================================
#
# This section handles automatic setup when the script is loaded.
# Currently disabled to improve script load performance, but can be
# re-enabled by uncommenting the initialization call below.
#
# WHAT IT DOES:
# - Initializes extension cache for all cloud environments
# - Pre-loads configuration data for faster lookups
# - Runs silently in background without user interaction
#
# TRADE-OFFS:
# - Faster subsequent extension lookups vs slower initial script load
# - Network calls during script initialization vs on-demand loading
# 
# To re-enable automatic initialization, uncomment the line below:
#==============================================================================

# Initialize extension cache silently when script loads (directly in main scope)
#Initialize-ExtensionCache | Out-Null

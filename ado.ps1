###
# Azure DevOps Work Item Helper Functions
###

function Get-AdoAuthHeader {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string] $Pat = $env:AZDO_PAT
    )

    if ([string]::IsNullOrWhiteSpace($Pat)) {
        throw "AZDO_PAT env var is empty. Set it first: `$env:AZDO_PAT = '<PAT>'"
    }

    $base64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat"))
    return @{ Authorization = "Basic $base64" }
}

function New-AdoWorkItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Title,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $Type = 'Bug',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $State = 'Active',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $Organization = 'msazure',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $Project = 'One',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $Area = 'One\Azure Portal\Hubs',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $Iteration = 'One\Krypton',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $ApiVersion = '7.1',

        [Parameter()]
        [switch] $QuietlyReturn
    )

    $uri = 'https://dev.azure.com/{0}/{1}/_apis/wit/workitems/${2}?api-version={3}' -f `
            $Organization, $Project, $Type, $ApiVersion

    $AssignedTo = "aksingal@microsoft.com"

    $bodyObj = @(
        @{ op = 'add'; path = '/fields/System.Title';         value = $Title }
        @{ op = 'add'; path = '/fields/System.AreaPath';      value = $Area }
        @{ op = 'add'; path = '/fields/System.IterationPath'; value = $Iteration }
        @{ op = 'add'; path = '/fields/System.AssignedTo';    value = $AssignedTo }
        @{ op = 'add'; path = '/fields/System.State';         value = $State }
    )

    $headers = Get-AdoAuthHeader
    $json    = $bodyObj | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers `
            -ContentType 'application/json-patch+json' -Body $json

        $retVal = "#$($response.id): $($response.fields.'System.Title')"

        if ($QuietlyReturn) {
            return $retVal
        } else {
            Write-Host "$($Type) $($retVal) ($($State))" -ForegroundColor Green
            Write-Host "https://msazure.visualstudio.com/One/_workitems/edit/$($response.id)" -ForegroundColor Green
        }
    }
    catch {
        # Make REST failures easier to debug
        $msg = $_.Exception.Message
        throw "Failed to create work item. URI=$uri. Error=$msg"
    }
}

###
# Convenience Wrappers - to create common work item types
###

function wi {
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

function bug {
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

    wi -Kind bug -Title $Title @PSBoundParameters
}

function pbi {
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

    wi -Kind pbi -Title $Title @PSBoundParameters
}

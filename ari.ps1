<#
.SYNOPSIS
Runs Azure Resource Inventory and uploads the generated files.

.DESCRIPTION
Generates an Azure Resource Inventory Excel report and Draw.io diagram, then uploads the files to the configured customer path unless upload is skipped.

.PARAMETER CustomerName
Customer name used in the upload path.

.PARAMETER SasToken
SAS token used for uploads.

.PARAMETER StorageBaseUrl
Base URL for the upload endpoint.

.PARAMETER ReportName
Report name passed to Azure Resource Inventory.

.PARAMETER SkipUpload
Skips the upload step and leaves the generated files on disk.

.EXAMPLE
pwsh ./ari.ps1

.EXAMPLE
pwsh ./ari.ps1 -CustomerName contoso -SasToken '?sv=...'

.EXAMPLE
pwsh ./ari.ps1 -SkipUpload
#>
[CmdletBinding()]
param(
	[string]$CustomerName = $env:ARI_CUSTOMER_NAME,
	[string]$SasToken = $env:ARI_SAS_TOKEN,
	[string]$StorageBaseUrl = "https://aridata.enabledapp.com",
	[string]$ReportName = "ARI",
	[switch]$SkipUpload
)

$ErrorActionPreference = "Stop"
$scriptTitle = "eGET-ARI"

if ($PSVersionTable.PSVersion.Major -lt 7) {
	throw "This script requires PowerShell 7 or later because AzureResourceInventory requires PowerShell 7+. Run it with 'pwsh ./ari.ps1'."
}

function Write-Step {
	param(
		[Parameter(Mandatory = $true)]
		[string]$Message
	)

	Write-Host "[$scriptTitle] $Message"
}

function Read-RequiredValue {
	param(
		[Parameter(Mandatory = $true)]
		[string]$Name,

		[Parameter(Mandatory = $true)]
		[string]$Prompt,

		[string]$CurrentValue
	)

	if (-not [string]::IsNullOrWhiteSpace($CurrentValue)) {
		return $CurrentValue
	}

	$enteredValue = Read-Host -Prompt $Prompt

	if ([string]::IsNullOrWhiteSpace($enteredValue)) {
		throw "$Name is required."
	}

	return $enteredValue
}

function Get-OutputDirectory {
	if ($IsWindows) {
		return "C:\AzureResourceInventory"
	}

	return Join-Path -Path $HOME -ChildPath "AzureResourceInventory"
}

function Reset-OutputDirectory {
	param(
		[Parameter(Mandatory = $true)]
		[string]$DirectoryPath
	)

	if (-not (Test-Path -LiteralPath $DirectoryPath)) {
		New-Item -Path $DirectoryPath -ItemType Directory -Force | Out-Null
		return
	}

	Get-ChildItem -LiteralPath $DirectoryPath -Force | Remove-Item -Recurse -Force
}

function Ensure-Module {
	param(
		[Parameter(Mandatory = $true)]
		[string]$Name
	)

	$available = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending

	if ($available) {
		return
	}

	try {
		Install-Module -Name $Name -Scope CurrentUser -Force -WarningAction SilentlyContinue -ErrorAction Stop
		return
	}
	catch {
		$moduleCacheRoot = Join-Path -Path $HOME -ChildPath ".ari\Modules"

		if (-not (Test-Path -LiteralPath $moduleCacheRoot)) {
			New-Item -Path $moduleCacheRoot -ItemType Directory -Force | Out-Null
		}

		Save-Module -Name $Name -Path $moduleCacheRoot -Force -WarningAction SilentlyContinue -ErrorAction Stop

		if ($env:PSModulePath -notlike "*$moduleCacheRoot*") {
			$env:PSModulePath = "$moduleCacheRoot;$env:PSModulePath"
		}
	}
}

function Remove-OldUserModuleVersions {
	param(
		[Parameter(Mandatory = $true)]
		[string]$Name,

		[Parameter(Mandatory = $true)]
		[Version]$KeepVersion,

		[string]$KeepModuleBase
	)

	$homePrefix = $HOME.TrimEnd('/\\')
	$oldVersions = Get-Module -ListAvailable -Name $Name |
		Where-Object {
			(($_.Version -lt $KeepVersion) -or ($KeepModuleBase -and $_.Version -eq $KeepVersion -and $_.ModuleBase -ne $KeepModuleBase)) -and
			$_.ModuleBase -like "$homePrefix*"
		} |
		Sort-Object Version -Descending

	foreach ($oldModule in $oldVersions) {
		try {
			Write-Step "Removing old user-installed module $Name $($oldModule.Version)"
			Uninstall-Module -Name $Name -RequiredVersion $oldModule.Version -Force -ErrorAction Stop
		}
		catch {
			Write-Step "Could not remove $Name $($oldModule.Version), continuing"
		}
	}
}

function Ensure-SingleModuleVersion {
	param(
		[Parameter(Mandatory = $true)]
		[string]$Name
	)

	$available = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending
	if (-not $available) {
		return
	}

	$keep = $available | Select-Object -First 1
	Remove-OldUserModuleVersions -Name $Name -KeepVersion $keep.Version -KeepModuleBase $keep.ModuleBase
}

function Get-LatestGeneratedFile {
	param(
		[Parameter(Mandatory = $true)]
		[string]$DirectoryPath,

		[Parameter(Mandatory = $true)]
		[string]$Filter
	)

	$generatedFile = Get-ChildItem -LiteralPath $DirectoryPath -Filter $Filter -File |
		Sort-Object -Property LastWriteTimeUtc -Descending |
		Select-Object -First 1

	if (-not $generatedFile) {
		throw "No file matched '$Filter' in '$DirectoryPath'."
	}

	return $generatedFile
}

function Get-UploadUri {
	param(
		[Parameter(Mandatory = $true)]
		[string]$BaseUrl,

		[Parameter(Mandatory = $true)]
		[string]$CustomerName,

		[Parameter(Mandatory = $true)]
		[string]$BlobName,

		[string]$Token
	)

	$normalizedBaseUrl = $BaseUrl.TrimEnd('/')
	$normalizedToken = if ([string]::IsNullOrWhiteSpace($Token)) {
		""
	}
	elseif ($Token.StartsWith('?')) {
		$Token
	}
	else {
		"?$Token"
	}

	$escapedCustomerName = [System.Uri]::EscapeDataString($CustomerName)
	$escapedBlobName = [System.Uri]::EscapeDataString($BlobName)

	return "$normalizedBaseUrl/$escapedCustomerName/$escapedBlobName$normalizedToken"
}

function Get-UploadErrorMessage {
	param(
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.ErrorRecord]$ErrorRecord,

		[Parameter(Mandatory = $true)]
		[string]$TargetName
	)

	$exception = $ErrorRecord.Exception
	$statusCode = $null
	$responseBody = $null

	if ($exception.PSObject.Properties.Match('Response').Count -gt 0 -and $exception.Response) {
		try {
			$statusCode = [int]$exception.Response.StatusCode
		}
		catch {
		}

		try {
			if ($exception.Response.PSObject.Properties.Match('Content').Count -gt 0 -and $exception.Response.Content) {
				$responseBody = $exception.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
			}
			elseif ($exception.Response.PSObject.Methods.Match('GetResponseStream').Count -gt 0) {
				$responseStream = $exception.Response.GetResponseStream()
				if ($responseStream) {
					$streamReader = New-Object System.IO.StreamReader($responseStream)
					try {
						$responseBody = $streamReader.ReadToEnd()
					}
					finally {
						$streamReader.Dispose()
						$responseStream.Dispose()
					}
				}
			}
		}
		catch {
		}
	}

	$statusText = if ($null -ne $statusCode) {
		"HTTP $statusCode"
	}
	else {
		"No HTTP status"
	}

	$details = @(
		"Upload failed for '$TargetName'.",
		$statusText,
		$exception.Message
	)

	if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
		$details += $responseBody.Trim()
	}

	$details += "Share this message with support for troubleshooting."

	return ($details -join ' ')
}

function Get-StorageAccessErrorMessage {
	param(
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.ErrorRecord]$ErrorRecord,

		[Parameter(Mandatory = $true)]
		[string]$TargetName
	)

	$exception = $ErrorRecord.Exception
	$statusCode = $null
	$responseBody = $null

	if ($exception.PSObject.Properties.Match('Response').Count -gt 0 -and $exception.Response) {
		try {
			$statusCode = [int]$exception.Response.StatusCode
		}
		catch {
		}

		try {
			if ($exception.Response.PSObject.Properties.Match('Content').Count -gt 0 -and $exception.Response.Content) {
				$responseBody = $exception.Response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
			}
			elseif ($exception.Response.PSObject.Methods.Match('GetResponseStream').Count -gt 0) {
				$responseStream = $exception.Response.GetResponseStream()
				if ($responseStream) {
					$streamReader = New-Object System.IO.StreamReader($responseStream)
					try {
						$responseBody = $streamReader.ReadToEnd()
					}
					finally {
						$streamReader.Dispose()
						$responseStream.Dispose()
					}
				}
			}
		}
		catch {
		}
	}

	$statusText = if ($null -ne $statusCode) {
		"HTTP $statusCode"
	}
	else {
		"No HTTP status"
	}

	$details = @(
		"Storage access check failed for '$TargetName'.",
		"This check performs a write probe blob upload to validate CustomerName and SasToken.",
		$statusText,
		$exception.Message
	)

	if ($statusCode -eq 404) {
		$details += "HTTP 404 usually means the target container path was not found for this request. Validate CustomerName and SasToken exactly as provided."
	}
	elseif ($statusCode -eq 403) {
		$details += "HTTP 403 usually means the SAS token is invalid, expired, not yet valid, or missing the required write/create permissions. Validate SasToken and time window."
	}

	if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
		$details += $responseBody.Trim()
	}

	return ($details -join ' ')
}

function Test-IsTimeoutError {
	param(
		[Parameter(Mandatory = $true)]
		[System.Management.Automation.ErrorRecord]$ErrorRecord
	)

	$exception = $ErrorRecord.Exception

	if ($exception -is [System.TimeoutException]) {
		return $true
	}

	if ($exception -is [System.Threading.Tasks.TaskCanceledException]) {
		return $true
	}

	if ($exception -is [System.Net.WebException] -and $exception.Status -eq [System.Net.WebExceptionStatus]::Timeout) {
		return $true
	}

	if ($exception.PSObject.Properties.Match('InnerException').Count -gt 0 -and $exception.InnerException -is [System.TimeoutException]) {
		return $true
	}

	if ($exception.PSObject.Properties.Match('InnerException').Count -gt 0 -and $exception.InnerException -is [System.Net.WebException] -and $exception.InnerException.Status -eq [System.Net.WebExceptionStatus]::Timeout) {
		return $true
	}

	if ($exception.PSObject.Properties.Match('Response').Count -gt 0 -and $exception.Response) {
		try {
			$statusCode = [int]$exception.Response.StatusCode
			if ($statusCode -in @(408, 504)) {
				return $true
			}
		}
		catch {
		}
	}

	if ($exception.PSObject.Properties.Match('Message').Count -gt 0 -and $exception.Message -match 'timed out|timeout') {
		return $true
	}

	return $false
}

function Invoke-UploadWithRetry {
	param(
		[Parameter(Mandatory = $true)]
		[string]$FilePath,

		[Parameter(Mandatory = $true)]
		[string]$BlobName,

		[Parameter(Mandatory = $true)]
		[string]$UploadUri,

		[Parameter(Mandatory = $true)]
		[hashtable]$Headers,

		[int]$MaxAttempts = 3,

		[int]$TimeoutSeconds = 300
	)

	for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
		try {
			$attemptSuffix = ""
			if ($MaxAttempts -gt 1 -and $attempt -gt 1) {
				$attemptSuffix = " (attempt $attempt/$MaxAttempts)"
			}

			Write-Step "Uploading $BlobName$attemptSuffix"
			Invoke-RestMethod -Method Put -Headers $Headers -InFile $FilePath -Uri $UploadUri -TimeoutSec $TimeoutSeconds | Out-Null
			return
		}
		catch {
			if ((Test-IsTimeoutError -ErrorRecord $_) -and $attempt -lt $MaxAttempts) {
				Write-Step "Upload timed out for $BlobName, retrying"
				Start-Sleep -Seconds 2
				continue
			}

			throw (Get-UploadErrorMessage -ErrorRecord $_ -TargetName $BlobName)
		}
	}
}

function Normalize-CustomerContainerName {
	param(
		[Parameter(Mandatory = $true)]
		[string]$CustomerName
	)

	$normalized = $CustomerName.Trim().ToLowerInvariant()
	$normalized = [System.Text.RegularExpressions.Regex]::Replace($normalized, '[^a-z0-9-]+', '-')
	$normalized = [System.Text.RegularExpressions.Regex]::Replace($normalized, '-+', '-')
	$normalized = $normalized.Trim('-')

	if ([string]::IsNullOrWhiteSpace($normalized)) {
		throw "Provided CustomerName '$CustomerName' is not valid after normalization. Validate CustomerName exactly as provided."
	}

	if (-not $normalized.StartsWith('cust-')) {
		$normalized = "cust-$normalized"
	}

	if ($normalized.Length -gt 63) {
		throw "Provided CustomerName '$CustomerName' is too long after normalization ('$normalized'). Validate CustomerName exactly as provided."
	}

	if ($normalized -notmatch '^[a-z0-9](?:[a-z0-9-]{1,61})[a-z0-9]$') {
		throw "Provided CustomerName '$CustomerName' is not a valid storage container name after normalization ('$normalized'). Validate CustomerName exactly as provided."
	}

	return $normalized
}

function Test-StorageAccess {
	param(
		[Parameter(Mandatory = $true)]
		[string]$BaseUrl,

		[Parameter(Mandatory = $true)]
		[string]$CustomerName,

		[string]$Token
	)

	$probeBlobName = "ari-write-test-$([DateTime]::UtcNow.ToString('yyyyMMddHHmmss')).txt"
	$probeUploadUri = Get-UploadUri -BaseUrl $BaseUrl -CustomerName $CustomerName -BlobName $probeBlobName -Token $Token
	$probeHeaders = @{
		"x-ms-blob-type" = "BlockBlob"
		"x-ms-version" = "2023-11-03"
	}

	try {
		$emptyBody = New-Object byte[] 0
		Invoke-RestMethod -Method Put -Headers $probeHeaders -Body $emptyBody -Uri $probeUploadUri -TimeoutSec 30 -ErrorAction Stop | Out-Null
	}
	catch {
		throw (Get-StorageAccessErrorMessage -ErrorRecord $_ -TargetName $CustomerName)
	}
}

function Start-AriUpload {
	param(
		[string]$CustomerName,
		[string]$SasToken,
		[Parameter(Mandatory = $true)]
		[string]$StorageBaseUrl,
		[Parameter(Mandatory = $true)]
		[string]$ReportName,
		[switch]$SkipUpload
	)

	Write-Step "Resolving inputs"
	$customerNameInput = if ($SkipUpload) {
		$CustomerName
	}
	else {
		Read-RequiredValue -Name "CustomerName" -Prompt "Customer name" -CurrentValue $CustomerName
	}
	$sasToken = if ($SkipUpload) {
		$SasToken
	}
	else {
		Read-RequiredValue -Name "SasToken" -Prompt "SAS token" -CurrentValue $SasToken
	}

	$outputDirectory = Get-OutputDirectory

	Write-Step "Preparing output directory"
	Reset-OutputDirectory -DirectoryPath $outputDirectory

	Write-Step "Ensuring modules"
	Ensure-Module -Name "AzureResourceInventory"
	if (-not (Get-Module -Name AzureResourceInventory)) {
		Write-Step "Importing AzureResourceInventory. If you're using Azure Cloud Shell, ignore any warnings about Autosize and Auto-fitting columns"
		Import-Module AzureResourceInventory -WarningAction SilentlyContinue
	}

	Ensure-SingleModuleVersion -Name "ImportExcel"
	Ensure-SingleModuleVersion -Name "Az.ResourceGraph"
	Ensure-SingleModuleVersion -Name "Az.Accounts"
	Ensure-SingleModuleVersion -Name "Az.Storage"
	Ensure-SingleModuleVersion -Name "Az.Compute"
	Ensure-SingleModuleVersion -Name "Az.Monitor"
	Ensure-SingleModuleVersion -Name "Az.CostManagement"

	$customerContainerName = $null
	if (-not $SkipUpload) {
		$customerContainerName = Normalize-CustomerContainerName -CustomerName $customerNameInput
		Write-Step "Checking storage access"
		Test-StorageAccess -BaseUrl $StorageBaseUrl -CustomerName $customerContainerName -Token $sasToken
	}

	Write-Step "Generating inventory"
	Invoke-ARI -Lite -SecurityCenter -IncludeTags -IncludeCosts -ReportName $reportName -WarningAction SilentlyContinue | Out-Null

	Write-Step "Getting files"
	$reportFile = Get-LatestGeneratedFile -DirectoryPath $outputDirectory -Filter "$reportName`_Report_*.xlsx"
	$diagramFile = Get-LatestGeneratedFile -DirectoryPath $outputDirectory -Filter "$reportName`_Diagram_*.xml"

	$reportFilePath = $reportFile.FullName
	$reportFileName = $reportFile.Name
	$diagramFilePath = $diagramFile.FullName
	$diagramFileName = $diagramFile.Name

	if ($SkipUpload) {
		Write-Step "Upload skipped"
		Write-Host "[$scriptTitle] Report: $reportFilePath"
		Write-Host "[$scriptTitle] Diagram: $diagramFilePath"
		return
	}

	$uploadHeaders = @{ "x-ms-blob-type" = "BlockBlob" }
	$reportUploadUri = Get-UploadUri -BaseUrl $StorageBaseUrl -CustomerName $customerContainerName -BlobName $reportFileName -Token $sasToken
	$diagramUploadUri = Get-UploadUri -BaseUrl $StorageBaseUrl -CustomerName $customerContainerName -BlobName $diagramFileName -Token $sasToken

	Invoke-UploadWithRetry -FilePath $reportFilePath -BlobName $reportFileName -UploadUri $reportUploadUri -Headers $uploadHeaders
	Invoke-UploadWithRetry -FilePath $diagramFilePath -BlobName $diagramFileName -UploadUri $diagramUploadUri -Headers $uploadHeaders

	Write-Step "Script completed"
	Write-Host "[$scriptTitle] Report: $reportFilePath"
	Write-Host "[$scriptTitle] Diagram: $diagramFilePath"
}

Start-AriUpload -CustomerName $CustomerName -SasToken $SasToken -StorageBaseUrl $StorageBaseUrl -ReportName $ReportName -SkipUpload:$SkipUpload

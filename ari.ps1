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

	if (Get-Module -ListAvailable -Name $Name) {
		return
	}

	Install-Module -Name $Name -Scope CurrentUser -Force -WarningAction SilentlyContinue
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
		$normalized = "customer"
	}

	if (-not $normalized.StartsWith('cust-')) {
		$normalized = "cust-$normalized"
	}

	if ($normalized.Length -gt 63) {
		$normalized = $normalized.Substring(0, 63)
	}

	$normalized = $normalized.Trim('-')

	if ([string]::IsNullOrWhiteSpace($normalized)) {
		$normalized = "cust-customer"
	}

	return $normalized
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
	Ensure-Module -Name "Az.CostManagement"

	Write-Step "Importing AzureResourceInventory. If you're using Azure CloudShell, ignore any warnings about Autosize and Auto-fitting columns"
	Import-Module AzureResourceInventory -WarningAction SilentlyContinue

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

	$customerContainerName = Normalize-CustomerContainerName -CustomerName $customerNameInput
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

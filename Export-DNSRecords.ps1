<#
.SYNOPSIS
    Exports MX, TXT, and DMARC DNS records for a list of domains.

.DESCRIPTION
    Reads a CSV file containing a "Domain" column and queries DNS for MX, TXT,
    and DMARC (_dmarc TXT) records for each domain. Results are exported as JSON.

.PARAMETER CsvPath
    Path to the input CSV file. Must contain a "Domain" column.

.PARAMETER OutputPath
    Path for the JSON output file. Defaults to "DNSRecordsExport.json" in the
    same directory as the CSV.

.PARAMETER DnsServer
    Optional DNS server to use for queries. Defaults to system resolver.

.EXAMPLE
    .\Export-DNSRecords.ps1 -CsvPath .\domains.csv

.EXAMPLE
    .\Export-DNSRecords.ps1 -CsvPath .\domains.csv -OutputPath .\results.json -DnsServer 8.8.8.8
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$DnsServer
)

function Resolve-DnsRecordSafe {
    param (
        [string]$Name,
        [string]$Type,
        [string]$Server
    )

    $params = @{ Name = $Name; Type = $Type; ErrorAction = 'SilentlyContinue' }
    if ($Server) { $params['Server'] = $Server }

    try {
        Resolve-DnsName @params
    }
    catch {
        $null
    }
}

# Set default output path next to the CSV
if (-not $OutputPath) {
    $csvDir  = Split-Path -Parent (Resolve-Path $CsvPath)
    $OutputPath = Join-Path $csvDir 'DNSRecordsExport.json'
}

$domains = Import-Csv -Path $CsvPath

if (-not ($domains | Get-Member -Name 'Domain' -MemberType NoteProperty)) {
    Write-Error "CSV file must contain a 'Domain' column."
    exit 1
}

$results = [System.Collections.Generic.List[object]]::new()
$total   = ($domains | Measure-Object).Count
$index   = 0

foreach ($row in $domains) {
    $index++
    $domain = $row.Domain.Trim()

    if ([string]::IsNullOrWhiteSpace($domain)) { continue }

    Write-Progress -Activity "Querying DNS records" `
                   -Status "[$index/$total] $domain" `
                   -PercentComplete (($index / $total) * 100)

    Write-Verbose "Processing: $domain"

    # MX records
    $mxRecords = Resolve-DnsRecordSafe -Name $domain -Type MX -Server $DnsServer |
        Where-Object { $_.Type -eq 'MX' } |
        ForEach-Object {
            [PSCustomObject]@{
                Preference = $_.Preference
                Exchange   = $_.NameExchange
            }
        }

    # TXT records (root domain)
    $txtRecords = Resolve-DnsRecordSafe -Name $domain -Type TXT -Server $DnsServer |
        Where-Object { $_.Type -eq 'TXT' } |
        ForEach-Object { $_.Strings -join '' }

    # DMARC record (_dmarc TXT)
    $dmarcRecords = Resolve-DnsRecordSafe -Name "_dmarc.$domain" -Type TXT -Server $DnsServer |
        Where-Object { $_.Type -eq 'TXT' } |
        ForEach-Object { $_.Strings -join '' }

    $results.Add([PSCustomObject]@{
        Domain  = $domain
        MX      = @($mxRecords)
        TXT     = @($txtRecords)
        DMARC   = @($dmarcRecords)
    })
}

Write-Progress -Activity "Querying DNS records" -Completed

$results | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputPath -Encoding UTF8

Write-Host "Done. $($results.Count) domain(s) processed." -ForegroundColor Green
Write-Host "Output: $OutputPath" -ForegroundColor Cyan

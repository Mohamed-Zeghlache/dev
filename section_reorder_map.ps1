# PowerShell Script to Reorder ADAudit.ps1 Sections
# This script provides the exact mapping for reordering audit sections

$SectionMap = @{
    # Current Order -> New Logical Order
    "Enhanced Domain Controller Health Assessment" = @{ CurrentLines = "1109-1235"; NewPosition = 1; Category = "Core Infrastructure" }
    "DC Performance Metrics" = @{ CurrentLines = "2864-2970"; NewPosition = 2; Category = "Core Infrastructure" }
    "AD Database Health" = @{ CurrentLines = "2972-3103"; NewPosition = 3; Category = "Core Infrastructure" }
    "FSMO Roles Verification" = @{ CurrentLines = "1523-1792"; NewPosition = 4; Category = "Core Infrastructure" }
    "Replication Health Assessment" = @{ CurrentLines = "1238-1321"; NewPosition = 5; Category = "Core Infrastructure" }
    "DNS Configuration Testing" = @{ CurrentLines = "1326-1506"; NewPosition = 6; Category = "Core Infrastructure" }

    "AD Sites & Services Audit" = @{ CurrentLines = "1828-1867"; NewPosition = 7; Category = "Network & Sites" }
    "Trust Relationships Analysis" = @{ CurrentLines = "1794-1826"; NewPosition = 8; Category = "Network & Sites" }

    "Security Analysis & Risk Assessment" = @{ CurrentLines = "1993-2131"; NewPosition = 9; Category = "Security & Access Control" }
    "Privileged Account Monitoring" = @{ CurrentLines = "2687-2862"; NewPosition = 10; Category = "Security & Access Control" }
    "Protocol Security Analysis (LLMNR, SMB, TLS, NTLM)" = @{ CurrentLines = "2511-2582"; NewPosition = 11; Category = "Security & Access Control" }
    "Protocols & Legacy Services Check" = @{ CurrentLines = "2189-2243"; NewPosition = 12; Category = "Security & Access Control" }
    "LAPS Check" = @{ CurrentLines = "2623-2680"; NewPosition = 13; Category = "Security & Access Control" }

    "User & Computer Accounts Analysis" = @{ CurrentLines = "1929-1991"; NewPosition = 14; Category = "Accounts & Groups" }
    "Security Group Nesting and Membership Analysis" = @{ CurrentLines = "2441-2507"; NewPosition = 15; Category = "Accounts & Groups" }
    "Group Managed Service Accounts (gMSA) Usage Analysis" = @{ CurrentLines = "2587-2618"; NewPosition = 16; Category = "Accounts & Groups" }
    "Orphaned SID Analysis" = @{ CurrentLines = "2246-2295"; NewPosition = 17; Category = "Accounts & Groups" }

    "OU Hierarchy Mapping" = @{ CurrentLines = "2137-2185"; NewPosition = 18; Category = "Structure & Organization" }
    "Schema Architecture Analysis" = @{ CurrentLines = "3105-3228"; NewPosition = 19; Category = "Structure & Organization" }
    "AD Object Protection Analysis" = @{ CurrentLines = "3230-3468"; NewPosition = 20; Category = "Structure & Organization" }

    "Group Policy Assessment" = @{ CurrentLines = "1869-1929"; NewPosition = 21; Category = "Policies & Management"; Note = "Missing #endregion" }
    "Event Log Analysis" = @{ CurrentLines = "2299-2439"; NewPosition = 22; Category = "Policies & Management" }

    "Azure AD Connect Health" = @{ CurrentLines = "3476-3562"; NewPosition = 23; Category = "Optional Cloud/External Integration" }
    "Conditional Access Policy Check" = @{ CurrentLines = "3566-3640"; NewPosition = 24; Category = "Optional Cloud/External Integration" }
    "PKI Infrastructure Analysis" = @{ CurrentLines = "3644-3778"; NewPosition = 25; Category = "Optional Cloud/External Integration" }
}

# Display the reordering plan
Write-Host "=== ADAudit.ps1 Section Reordering Plan ===" -ForegroundColor Cyan
Write-Host ""

$Categories = @(
    "Core Infrastructure",
    "Network & Sites",
    "Security & Access Control",
    "Accounts & Groups",
    "Structure & Organization",
    "Policies & Management",
    "Optional Cloud/External Integration"
)

foreach ($Category in $Categories) {
    Write-Host "📁 $Category" -ForegroundColor Yellow
    $SectionsInCategory = $SectionMap.GetEnumerator() | Where-Object { $_.Value.Category -eq $Category } | Sort-Object { $_.Value.NewPosition }

    foreach ($Section in $SectionsInCategory) {
        $Note = if ($Section.Value.Note) { " ⚠️ $($Section.Value.Note)" } else { "" }
        Write-Host "  $($Section.Value.NewPosition). $($Section.Key) (Lines: $($Section.Value.CurrentLines))$Note" -ForegroundColor White
    }
    Write-Host ""
}

Write-Host "=== Next Steps ===" -ForegroundColor Green
Write-Host "1. Extract each section from the original file using the line ranges above"
Write-Host "2. Reorder them according to the new positions"
Write-Host "3. Add missing #endregion for Group Policy Assessment section"
Write-Host "4. Update progress tracking in web interface functions"
Write-Host "5. Test the reordered script"

# Function to extract a section from the original file
function Extract-Section {
    param(
        [string]$FilePath,
        [string]$StartLine,
        [string]$EndLine
    )

    $Start = [int]$StartLine
    $End = [int]$EndLine
    $Length = $End - $Start + 1

    Get-Content $FilePath | Select-Object -Skip ($Start - 1) -First $Length
}

# Example usage:
# $Section1 = Extract-Section -FilePath "ADAudit.ps1" -StartLine "1109" -EndLine "1235"
param
(
  [Parameter(Mandatory=$false,helpmessage="LEVEL")]
  [ValidateNotNullOrEmpty()]
  [string]$level="intermediate",
  [Parameter(Mandatory=$false,helpmessage="Filename")]
  [ValidateNotNullOrEmpty()]
  [string]$Filename="for INTERVIEWER DevOps-Jay Kim .xlsx",
  [Parameter(Mandatory=$false,helpmessage="Path")]
  [ValidateNotNullOrEmpty()]
  [string]$Path="/Users/Shared/Data/Documents/SoftServe/TI/CP/InProgress"
)
$ErrorActionPreference = "Stop"
Write-Host "CP Path $Path/$Filename"
$excel = Open-ExcelPackage "$Path/$Filename"

$Activities = $excel.Workbook.Worksheets['Competencies']
$Summary = $excel.Workbook.Worksheets['Summary']

"The candidate has been working in IT for about " + $Summary.Cells['B6'].Value + " years."
Write-Host "His main responsibilities were "
Write-Host "His strong areas: "
Write-Host ""
Write-Host "Candidate has strong experience in:"
for ($i=5; $i -lt 40; $i++)
{
    if ($Activities.Cells["D$i"].Value -eq "Strong")
    {
        if ($Activities.Cells["E$i"].Value)
        {
            "- " + $Activities.Cells["A$i"].Value + " (" + $Activities.Cells["E$i"].Value + ")"
        }
        else {
            "- " + $Activities.Cells["A$i"].Value
        } 
    }
}
Write-Host ""
Write-Host "He is also good at:"
for ($i=5; $i -lt 40; $i++)
{
    if ($Activities.Cells["D$i"].Value -eq "Good")
    {
        if ($Activities.Cells["E$i"].Value)
        {
            "- " + $Activities.Cells["A$i"].Value + " (" + $Activities.Cells["E$i"].Value + ")"
        }
        else {
            "- " + $Activities.Cells["A$i"].Value
        } 
    }
}
Write-Host ""
Write-Host "Candidate has basic understanding of:"
for ($i=5; $i -lt 40; $i++)
{
    if ($Activities.Cells["D$i"].Value -eq "Beginner")
    {
        if ($Activities.Cells["E$i"].Value)
        {
            "- " + $Activities.Cells["A$i"].Value + " (" + $Activities.Cells["E$i"].Value + ")"
        }
        else {
            "- " + $Activities.Cells["A$i"].Value
        } 
    }
}

Write-Host ""
Write-Host "Candidate has critical gaps in:"
for ($i=5; $i -lt 40; $i++)
{
    if ($Activities.Cells["D$i"].Value -eq "None")
    {
        if ($Activities.Cells["E$i"].Value)
        {
            "- " + $Activities.Cells["A$i"].Value + " (" + $Activities.Cells["E$i"].Value + ")"
        }
        else {
            "- " + $Activities.Cells["A$i"].Value
        }
    }
}

Write-Host ""
Write-Host "Leading/Mentorship experience:" $Summary.Cells["B11"].Value
Write-Host "Customer communication experience:" $Summary.Cells["B12"].Value
Write-Host "Technical English:    " $Summary.Cells["C19"].Value
Write-Host "Recommended projects/Roles: "
Write-Host ""
"Candidate has a " + $Summary.Cells["C17"].Value + " technical level according to KM."
Write-Host ""
Write-Host "Candidate has knowledge gaps according to KM in the following areas"
Switch ($level) {
    "trainee" {$cell="F"}    
    "junior" {$cell="G"}
    "intermediate" {$cell="H"}
    "senior" {$cell="I"}
    "lead" {$cell="J"}
}

for ($i=5; $i -lt 40; $i++)
{
    if ($Activities.Cells["$cell$i"].Value -eq "--")
    {
        "- " + $Activities.Cells["A$i"].Value
    }
}
Write-Host ""
#$wb.Close($false)
#$excel.Quit()
Write-Host "Close-ExcelPackage"
Close-ExcelPackage $excel
Write-Host "GC Collect"
[System.GC]::Collect()
#[System.GC]::WaitForPendingFinalizers()
Exit

#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Activities)
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Summary)
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

Remove-Variable -Name excel
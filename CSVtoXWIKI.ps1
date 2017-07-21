Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

Function Save-File($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    $SaveFileDialog.filter = "All files (*.*)| *.*"
    $SaveFileDialog.ShowDialog() | Out-Null
    $SaveFileDialog.filename
}

Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

Write-Host "Get the CSV file (comma delimited)"
$csv = Get-FileName "C:\"
$delimiter = ","
Write-Host "Where do you want to save your xwiki text file?"
$textdoc = Save-File "C:\"

$data = Import-Csv $csv

$numberOfColumns = (Get-Content $csv | 
    ForEach-Object{($_.split($delimiter)).Count} | 
    Measure-Object -Maximum | 
    Select-Object -ExpandProperty Maximum)

$numberOfRows=($data | Measure-Object | Select-Object).Count

foreach ($_ in (0..($numberOfColumns - 1)))
{
    New-Variable -Name "header$_" -Value ($data | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name')[$_]
}

New-Item $textdoc -type file -Force
Add-Content -Path $textdoc -NoNewline -Value "(:table border=0 cellpadding=5 cellspacing=0 class=listing:) "
ForEach ($_ in (0..($numberOfColumns - 1)))
{
    $var = Get-Variable -Name "header$_" -ValueOnly
    Add-Content -Path $textdoc -NoNewline -Value "(:cellnr class=listing-header:) $var "
}

Add-Content -Path $textdoc -Value "`n"

ForEach ($row in (0..($numberOfRows - 1)))
{
    ForEach ($_ in (0..($numberOfColumns - 1)))
    {
        $var = Get-Variable -Name "header$_" -ValueOnly
        $value = $data[$row].$var
        if ($var -eq "$header0")
        {
            Add-Content -Path $textdoc -Value "(:cellnr:) $value"
        }
        else
        {
            Add-Content -Path $textdoc -Value "(:cell:) $value"
        }
    }
}

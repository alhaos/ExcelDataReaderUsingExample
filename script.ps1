Add-Type -path ExcelDataReader.dll
$stream = [System.IO.File]::Open("Книга.xlsx", [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)

class Row {
    [string]$Company
    [string]$Contact
    [string]$Country
}

$Rows = @()
$edr = [ExcelDataReader.ExcelReaderFactory]::CreateReader($stream)

while ($edr.Read()) {
    $Rows += , [Row]@{
        Company = $edr.GetString(0)
        Contact = $edr.GetString(1)
        Country = $edr.GetString(2)  
    }
}

$edr.Close()

$Rows | Export-Csv -Encoding utf8 out.csv

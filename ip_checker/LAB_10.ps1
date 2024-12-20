# Określenie ścieżek do plików (domyślnie w tym samym folderze co skrypt)
$inputFilePath = "C:\Users\User\adresy.xlsx"
$outputFilePath = "C:\Users\User\wyniki.xlsx"

$ipAddresses = Import-Excel -Path $inputFilePath -WorksheetName 'IP-Addresses' |
               Select-Object -First 5 -ExpandProperty 'Adres IP'

if ($ipAddresses -eq $null -or $ipAddresses.Count -eq 0) {
    Write-Host "Brak adresow IP w pliku." -ForegroundColor Red
    exit
}

$results = @()

foreach ($ip in $ipAddresses) {
    if ($ip -match '^(\d{1,3}\.){3}\d{1,3}$') {
        $pingResult = Test-Connection -ComputerName $ip -Count 1 -Quiet

        $results += [PSCustomObject]@{
            Adres = $ip
            Wynik = if ($pingResult) { "Dostepny" } else { "Niedostepny" }
        }
    } else {
        $results += [PSCustomObject]@{
            Adres = $ip
            Wynik = "Niepoprawny adres IP"
        }
    }
}

$results | Export-Excel -Path $outputFilePath -WorksheetName 'Wyniki' -AutoSize

Write-Host "Wyniki zostasly zapisane do $outputFilePath" -ForegroundColor GreenS

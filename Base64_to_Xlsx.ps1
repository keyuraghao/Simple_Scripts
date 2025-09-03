$base64String = ""
$outputFilePath = "output.xlsx"

try {
    $decodedBytes = [System.Convert]::FromBase64String($base64String)
    Set-Content -Path $outputFilePath -Value $decodedBytes -Encoding Byte

    Write-Host "Excel file successfully decoded and saved to: $outputFilePath"
} catch {
    Write-Host "Error decoding base64 and saving Excel file: $_"
}
Apollo.io

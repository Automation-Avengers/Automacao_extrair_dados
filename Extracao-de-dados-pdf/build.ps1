$exclude = @("venv", "Extracao-de-dados-pdf.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "Extracao-de-dados-pdf.zip" -Force
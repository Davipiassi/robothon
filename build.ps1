$exclude = @("venv", "CardsOperations.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "CardsOperations.zip" -Force
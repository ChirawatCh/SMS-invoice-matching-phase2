# List of executable files to run in order
$executables = @(
    "cpaxtra_v2.exe",
    "cpram_v2.exe",
    "perfect-companion_v2.exe",
    "president_V2.exe",
    "thainamthip_V2.exe",
    "concat_csv.exe"
)

# Loop through the executables and run each one
foreach ($exe in $executables) {
    Write-Host "Running $exe"
    Start-Process -FilePath $exe -Wait
}

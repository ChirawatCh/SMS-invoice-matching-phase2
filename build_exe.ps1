# Save the current directory
$currentDir = Get-Location

# Change directory to where the .py files are located
Set-Location -Path "scripts"

# Run PyInstaller for each script
& pyinstaller --onefile cpaxtra_v2.py
& pyinstaller --onefile cpram_v2.py
& pyinstaller --onefile perfect-companion_v2.py
& pyinstaller --onefile president_V2.py
& pyinstaller --onefile thainamthip_V2.py
& pyinstaller --onefile concat_csv.py

# Remove '*.spec' files
Remove-Item -Path *.spec -Force

# Move '*.exe' files to the parent folder
Move-Item -Path ".\dist\*" -Destination "..\" -Force

# Remove 'dist' folder
Remove-Item -Path ".\dist" -Force -Recurse

# Return to the original directory
Set-Location -Path $currentDir

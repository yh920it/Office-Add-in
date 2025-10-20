# Office-Add-in
Office add-in to automate follow-up reminder

# --- CONFIGURE THESE ---
$ezPath = "C:\Users\Public\EZStation"            # <-- set your EZStation folder path
$backupRoot = "C:\Temp\EZStationBackups"        # <-- where backups will be stored
$processName = "EZStation"                      # <-- process or service name (adjust)
# ------------------------

# Timestamped vars
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = Join-Path $backupRoot "EZStationBackup_$ts"
$quarantinePath = Join-Path $backupPath "Quarantine_AdminState"

# 1) create backup dirs
New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
New-Item -Path $quarantinePath -ItemType Directory -Force | Out-Null

# 2) stop process (best-effort)
Write-Output "Stopping process '$processName' (if present)..."
Stop-Process -Name $processName -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# 3) full backup (copy everything)
Write-Output "Backing up EZStation folder from $ezPath to $backupPath ..."
Copy-Item -Path "$ezPath\*" -Destination $backupPath -Recurse -Force
Write-Output "Backup complete."

# 4) Inspect files and show a preview to the operator
Write-Output "Previewing files in $ezPath (name, size, lastwrite):"
Get-ChildItem -Path $ezPath -File | Select-Object Name, Length, LastWriteTime | Sort-Object LastWriteTime -Descending | Format-Table -AutoSize

# 5) Move only likely admin/user state files to quarantine
#    Patterns below are conservative — adjust as needed. They avoid .db and large files.
$patterns = @("user","account","admin","passwd","password",".usr",".profile",".cfg",".config",".ini","login")

foreach ($p in $patterns) {
    Get-ChildItem -Path $ezPath -Filter $p -File -ErrorAction SilentlyContinue | ForEach-Object {
        $src = $_.FullName
        $dest = Join-Path $quarantinePath $_.Name
        Move-Item -Path $src -Destination $dest -Force
        Write-Output ("Quarantined: " + $_.Name)
    }
}


# 6) Show what we quarantined
Write-Output "Quarantine contents:"
Get-ChildItem -Path $quarantinePath -File | Select-Object Name, Length, LastWriteTime | Format-Table -AutoSize

# 7) Start EZStation process or prompt user to manually start it
Write-Output "You should now start EZStation (manually or via service). The app should prompt to create a new Super Admin if these were the right files."
# Example to start exe (uncomment and set path if you know it):
# Start-Process -FilePath "C:\Program Files\EZStation\EZStation.exe"

# --- Rollback helpers (won't run automatically) ---
Write-Output "To restore everything (rollback), run the following commands:"
Write-Output "1) Stop EZStation service/process."
Write-Output "2) Remove or move the current EZStation folder:"
Write-Output "   Remove-Item -Path '$ezPath' -Recurse -Force"
Write-Output "3) Restore from backup:"
Write-Output "   Copy-Item -Path '$backupPath\*' -Destination '$ezPath' -Recurse -Force"
Write-Output "4) Start EZStation."

Write-Output "Operations complete. Please test device list, live view, and recordings now."

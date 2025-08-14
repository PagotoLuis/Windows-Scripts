netstat -b | ForEach-Object { 
    if ($_ -match "chrome.exe") { 
        Write-Host $_; 
        $next = $true 
    } 
    elseif ($next) { 
        Write-Host $_; 
        $next = $false 
    } 
}

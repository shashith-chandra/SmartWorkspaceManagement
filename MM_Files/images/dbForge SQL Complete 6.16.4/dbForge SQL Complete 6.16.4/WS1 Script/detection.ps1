if (test-path "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{99999999-9999-9999-9999-999999999999}"){
    if ($(Get-ItemProperty -Path "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{99999999-9999-9999-9999-999999999999}" -Name "DisplayVersion").DisplayVersion -eq "1.0"){
        Exit 0
    } 
}

Exit 9990
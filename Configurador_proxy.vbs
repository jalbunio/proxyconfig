dim ConfProxy
set ConfProxy = Wscript.CreateObject("Wscript.Shell")

if msgbox("Habilitar Proxy? Se está fora do escritório, clique em 'Nao'.", vbQuestion or vbYesNo) = vbYes then
ConfProxy.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD" 
ConfProxy.RegWrite "HKCU\Software\Microsoft\Windows\currentVersion\Internet Settings\ProxyServer", "192.168.0.1:3128", "REG_SZ" 
ConfProxy.RegWrite "HKCU\Software\Microsoft\Windows\currentVersion\Internet Settings\ProxyOverride", "*.getbootstrap.com;<local>", "REG_SZ"
else 

ConfProxy.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD" 
End if 
Set ConfProxy = Nothing
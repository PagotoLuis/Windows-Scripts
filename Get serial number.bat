cd C:\DOCUMENTOS\PROGRAMAS\PSTools


set /p Input=Digite o nome da Maquina:
copy C:\Documentos\PROGRAMAS\agent_metasa_tz0.msi \\%Input%\c$\agent_metasa_tz0.msi

psexec \\%Input% powershell
Get-WmiObject win32_bios | select Serialnumber
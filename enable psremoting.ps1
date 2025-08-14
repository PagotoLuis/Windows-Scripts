
Set-Location C:\DOCUMENTOS\PROGRAMAS
Set-Location PSTools

$MACHINE = Read-Host "Digite o nome da maquina: "

.\PsExec.exe \\$MACHINE powershell -Command Enable-PSRemoting

exit
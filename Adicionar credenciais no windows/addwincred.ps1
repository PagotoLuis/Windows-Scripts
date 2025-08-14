$targetName = "UMA039S"
$targetName1 = "UMA015S"
$targetName2 = "UMA004S"
$targetName3 = "UMA042S"


$username = Read-Host "Usuario"
$passwordSecure = Read-Host "Senha" -AsSecureString


$passwordBSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordSecure)
$passwordPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($passwordBSTR)


Write-Host "Adicionando credenciais ao Gerenciador de Credenciais..." -ForegroundColor Cyan
try {
    cmdkey /add:$targetName /user:$username /pass:$passwordPlain
    cmdkey /add:$targetName1 /user:$username /pass:$passwordPlain
    cmdkey /add:$targetName2 /user:$username /pass:$passwordPlain
    cmdkey /add:$targetName3 /user:$username /pass:$passwordPlain
    Write-Host "Credenciais do Windows adicionadas com sucesso!" -ForegroundColor Green
} catch {
    Write-Host "Erro ao adicionar credenciais: $_" -ForegroundColor Red
} finally {
    # Limpar a senha da memória
    $passwordPlain = null
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordBSTR)
}

Read-Host "Pressione enter para continuar..."

# Invoke-PS2EXE -InputFile "defaultMachines.ps1" -OutputFile ".\defaultMachines.exe" -company "Metasa" -version "3.1" -iconFile "C:\Users\luis.pagoto\Pictures\ico.ico"
# Verificando se o Powershell esta sendo executado como Administador
$admin = [System.Security.Principal.WindowsPrincipal] [System.Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $admin.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Este script precisa ser executado como administrador" -ForegroundColor Red
    Pause
    Exit
}

# Para iniciar o loading: execute `runLoading`
# Para interromper: execute `stopLoading`
$servidor = ''
$servidor2 = ''
$servidor3 = ''
$servidor4 = ''

#LEITURA DO TERMINAL -----------------------------------------
function printMenu {
    Clear-Host
    Write-Host "       ________________________________________________________________________ " -ForegroundColor Blue
    Write-Host "      |                                                                        |" -ForegroundColor Blue
    Write-Host "      |               Instalador Automatico de Programas                       |" -ForegroundColor Blue
    Write-Host "      |                                                                        |" -ForegroundColor Blue
    Write-Host "      |___________________________Luis Pagoto__________________________________|" -ForegroundColor Blue
    Write-Host "      "
    Write-Host "      [  1  ] . Instalacao inicial" -ForegroundColor Blue
    Write-Host "      [  2  ] . Instalacao padrao" -ForegroundColor Blue
    Write-Host "      [  3  ] . Instalacao para engenharia" -ForegroundColor Blue
    Write-Host "      [  4  ] . Instalacao para computadores remotos" -ForegroundColor Blue
    Write-Host "      [  5  ] . Configuracoes e otimizacoes" -ForegroundColor Blue
    Write-Host "      [  6  ] . Selecionar um especifico" -ForegroundColor Blue
    Write-Host "      [  7  ] . Gerador de senha" -ForegroundColor Blue
    Write-Host "      [  8  ] . Coletar Informação de uma máquina" -ForegroundColor Blue
    Write-Host "      [  9  ] . Fazer backup" -ForegroundColor Blue
    Write-Host "      [  0  ] . Sair" -ForegroundColor Blue
    Write-Host ""
    return Read-Host "      Opcao [0-9]"
}

function readSelect {
    do{
        Write-Host "      [  1   ] . Instalar TraumaZero" -ForegroundColor Blue
        Write-Host "      [  2   ] . Instalar Antivirus" -ForegroundColor Blue
        Write-Host "      [  3   ] . Instalar Winget" -ForegroundColor Blue
        Write-Host "      [  4   ] . Instalar Programas Basicos" -ForegroundColor Blue
        Write-Host "      [  5   ] . Instalar Office 365" -ForegroundColor Blue
        Write-Host "      [  6   ] . Instalar Datasul" -ForegroundColor Blue
        Write-Host "      [  7   ] . Instalar RM" -ForegroundColor Blue
        Write-Host "      [  8   ] . Instalar Oracle" -ForegroundColor Blue
        Write-Host "      [  9   ] . Instalar Sapiens" -ForegroundColor Blue
        Write-Host "      [  10  ] . Instalar VetorH" -ForegroundColor Blue
        Write-Host "      [  11  ] . Verificar Atualizacoes" -ForegroundColor Blue
        Write-Host "      [  12  ] . Realizar limpeza de disco" -ForegroundColor Blue
        Write-Host "      [  13  ] . Ajustar plano de energia p/ 10 min monitor e suspensao para nunca" -ForegroundColor Blue
        Write-Host "      [  14  ] . Copiar link Remoto Marau" -ForegroundColor Blue
        Write-Host "      [  15  ] . Instalar Anydesk" -ForegroundColor Blue
        Write-Host "      [  16  ] . Instalar Tekla 2025" -ForegroundColor Blue
        Write-Host "      [  17  ] . Instalar Tekla 2024" -ForegroundColor Blue
        Write-Host "      [  18  ] . Instalar Tekla 2022" -ForegroundColor Blue
        Write-Host "      [  19  ] . Instalar Trimble Connect" -ForegroundColor Blue
        Write-Host "      [  20  ] . Instalar Autocad 2018" -ForegroundColor Blue
        Write-Host "      [  21  ] . Instalar Navisworks Freedom" -ForegroundColor Blue
        Write-Host "      [  22  ] . Instalar Mfiles" -ForegroundColor Blue
        Write-Host "      [  23  ] . Instalar True View" -ForegroundColor Blue
        Write-Host "      [  0   ] . Voltar" -ForegroundColor Blue
        $opSelect = Read-Host "      Opcao [0-23] : "
        switch ($opSelect) {
            1 {RunTrauma "s"}
            2 {RunSymantec "s" }
            3 {RunWinget "s"}
            4 {RunPrograms "s"}
            5 {RunOffice "S" 1}
            6 {RunTotvs "s" "n"}
            7 {RunTotvs "n" "s"}
            8 {RunOracle "s"}
            9 {RunSapiens "s"}
            10 {RunVetorH "s"} 
            11 {RunUpdate "s"}
            12 {RunCleanup "s"}
            13 {RunEnergy "s"}
            14 {RunRdp "s"}
            15 {winget install -e --id AnyDeskSoftwareGmbH.AnyDesk --accept-package-agreements}
            16 {RunTekla "n" "n" "s"}
            17 {RunTekla "n" "s" "n"}
            18 {RunTekla "s" "n" "n"}
            19 {RunTrimble "s"}
            20 {RunAutocad "s"}
            21 {RunNavis "s"}
            22 {RunMfiles "s"}
            23 {RunTruview "s"} 
            24 {Write-Host "`nSaindo..." -ForegroundColor Red}
            Default {Write-Host "`nOpcao invalida Por favor, tente novamente." -ForegroundColor Red}
        }
    } while ($opSelect -ne 0)
}

function readStandard {
    Write-Host "----------------  Instalacao padrao ----------------------------------------" -ForegroundColor Blue

    $configuracoes = [PSCustomObject]@{
        TraumaZero      = Read-Host "      Instalar TraumaZero (S/N)?"
        Antivirus       = Read-Host "      Instalar Antivirus (S/N)?"
        Winget          = Read-Host "      Instalar Winget (S/N)?"
        SoftwaresBasicos = Read-Host "      Instalar Programas Basicos (S/N)?"
        Office          = Read-Host "      Instalar Office (S/N)?"
        OfficeVersao    = $null
        Datasul         = Read-Host "      Instalar Datasul (S/N)?"
        RM              = Read-Host "      Instalar RM (S/N)?"
        Oracle          = Read-Host "      Instalar Oracle (S/N)?"
        Sapiens         = Read-Host "      Instalar Sapiens (S/N)?"
        VetorH          = Read-Host "      Instalar VetorH (S/N)?"
    }

    if ($configuracoes.Office -match "^[SsYy]$") {
        Write-Host "      `n1 - Office 365"
        Write-Host "      2 - Office 2013"
        Write-Host "      3 - Office 2019"
        Write-Host "      4 - Office 2021`n"

        while ($true) {
            $configuracoes.OfficeVersao = Read-Host "      Escolha a versao a ser instalada (1-4)"
            if ($configuracoes.OfficeVersao -match "^[1-4]$") {
                break
            } else {
                Write-Host "      Valor invalido Por favor, escolha uma versao entre 1 e 4." -ForegroundColor Red
            }
        }
    }

    return $configuracoes
}

function readConfig {
    Write-Host "---------------- Configuracoes para otimizacoes ------------------------" -ForegroundColor Blue

    $configuracoes = [PSCustomObject]@{
        Update      = Read-Host "      Verificar Atualizacoes (S/N)?"
        Cleandisk   = Read-Host "      Realizar limpeza de disco (S/N)?"
        Energy      = Read-Host "      Ajustar plano de energia p/ 10 min monitor e suspensao para nunca (S/N)?"
    }

    return $configuracoes
}

function readRemote {
    Write-Host "----------------  Instalacao para computadores remotos ---------------------" -ForegroundColor Blue

    $configuracoes = [PSCustomObject]@{
        Rdp      = Read-Host "      Copiar link Remoto Marau (S/N)?"
        Anydesk  = Read-Host "      Instalar Anydesk (S/N)?"
    }

    return $configuracoes
}

function readEng {
    Write-Host "----------------  Instalacao para computadores da engenharia ---------------" -ForegroundColor Blue

    $configuracoes = [PSCustomObject]@{
        Tekla25  = Read-Host "      Instalar Tekla 2025 (S/N)?"
        Tekla24  = Read-Host "      Instalar Tekla 2024 (S/N)?"
        Tekla22  = Read-Host "      Instalar Tekla 2022 (S/N)?"
        Trimble  = Read-Host "      Instalar Trimble Connect (S/N)?"
        Autocad18 = Read-Host "      Instalar Autocad 2018 (S/N)?"
        Navis    = Read-Host "      Instalar Navisworks Freedom (S/N)?"
        Mfiles   = Read-Host "      Instalar Mfiles (S/N)?"
        Trueview = Read-Host "      Instalar True View (S/N)?"
    }

    return $configuracoes
}

#EXECUcao DAS FUNcoes --------------------------------------------
        #------------------------------Configurar opcoes de energia
function RunEnergy($isEnergy) {
    
    if($isEnergy -eq "S" -or $isEnergy -eq "s" -or $isEnergy -eq "Y" -or $isEnergy -eq "y"){
        Write-Output "Alterado configuracao de desligamento de tela para nunca"
        powercfg -change standby-timeout-ac 0
        powercfg -change monitor-timeout-ac 10
        powercfg -change standby-timeout-dc 0
        powercfg -change monitor-timeout-dc 10
    }
}
        #------------------------------Instalar RDP
function RunRdp($isRdp){
    if($isRdp -eq "S" -or $isRdp -eq "s" -or $isRdp -eq "Y" -or $isRdp -eq "y"){
        Copy-Item \\$servidor\softwares$\RemotoMarau.rdp C:\users\Public\desktop
        Write-Output "Link Remoto Marau copiado para a area de trabalho"
    }
}

        #------------------------------Instalar Winget
function RunWinget($isWinget){
    if($isWinget -eq "S" -or $isWinget -eq "s" -or $isWinget -eq "Y" -or $isWinget -eq "y"){
        Write-Output "Instalando o Winget..."
        Start-Process -FilePath "\\$servidor\softwares$\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    }
}

        #------------------------------Instalar Trauma Zero
function RunTrauma($isTrauma){
    if($isTrauma -eq "S" -or $isTrauma -eq "s" -or $isTrauma -eq "Y" -or $isTrauma -eq "y"){
        Write-Output "Instalando o TraumaZero"
        Start-Process -FilePath "\\metasa\NETLOGON\agent_metasa_tz0.msi" -Wait
    }
}

        #------------------------------Instalar Oracle
function RunOracle($isOracle){
    if($isOracle -eq "S" -or $isOracle -eq "s" -or $isOracle -eq "Y" -or $isOracle -eq "y"){
        Write-Output "Instalando o Oracle Client 19..."
        Start-Process -FilePath "\\$servidor\softwares$\Oracle\client19c_x86\setup.exe" -Wait
        Copy-Item \\$servidor\softwares$\Oracle\tnsnames.ora C:\Oracle\product\19.0.0\client_1\network\admin
        Copy-Item \\$servidor\softwares$\Oracle\tnsnames.ora C:\Oracle\product\19.0.0\client_1\network\admin\Sample
        Write-Output "Arquivos de configuracao Oracle copiados"
    }
}

        #------------------------------Instalar Sapiens
function RunSapiens($isSapiens){
    if($isSapiens -eq "S" -or $isSapiens -eq "s" -or $isSapiens -eq "Y" -or $isSapiens -eq "y"){
        Write-Output "Instalando o Sapiens..."
        Start-Process -FilePath  "\\$servidor2\Senior\SeniorInstaller.exe" -Wait
    }
}
        #------------------------------Instalar VetorH
function RunVetorH($isVetorh){
    if($isVetorh -eq "S" -or $isVetorh -eq "s" -or $isVetorh -eq "Y" -or $isVetorh -eq "y"){
        Write-Output "Instalando o VetorH..."
        Start-Process -FilePath  "\\$servidor4\Senior\SeniorInstaller.exe" -Wait
    }
}

        #------------------------------Instalar TOTVS Datasul
function RunTotvs($isDatasul, $isRM){
    if($isDatasul -eq "S" -or $isDatasul -eq "s" -or $isDatasul -eq "Y" -or $isDatasul -eq "y"){
        Write-Output "Instalando o Datasul..."
        Copy-Item "\\$servidor\softwares$\TOTVSSA\TOTVS\GraphOn" -Destination "C:\TOTVS\GraphOn" -Recurse
        Copy-Item "\\$servidor\softwares$\TOTVSSA\Datasul.lnk" -Destination "C:\Users\Public\Desktop\"
    }
    #------------------------------Instalar TOTVS RM
    if($isRM -eq "S" -or $isRM -eq "s" -or $isRM -eq "Y" -or $isRM -eq "y"){
        Write-Output "Instalando o RM..."
        Copy-Item "\\$servidor\softwares$\TOTVSSA\TOTVS\RM.NET" -Destination "C:\TOTVS\" -Recurse
        Copy-Item "\\$servidor\softwares$\TOTVSSA\RM.lnk" -Destination "C:\Users\Public\Desktop\"
    }
}

        #------------------------------Instalar Antivirus
function RunSymantec($isAV){
    if($isAV -eq "S" -or $isAV -eq "s" -or $isAV -eq "Y" -or $isAV -eq "y"){
        Write-Output "Instalando o Antivirus $isAVDesc..."
        Start-Process -FilePath "\\$servidor3\SES_Cloud_Clients\Marau\Marau_Estacoes_Novas_Symantec_Agent_setup_EN_RU9.exe" -Wait
    }
}

        #------------------------------Instalar Programas powershell admin
function RunPrograms($isSoftwares, $isAnydesk){
    if($isSoftwares -eq "S" -or $isSoftwares -eq "s" -or $isSoftwares -eq "Y" -or $isSoftwares -eq "y"){
        Write-Output "Instalando os programas..."
        winget install -e --id Google.Chrome --accept-package-agreements --ignore-security-hash
        winget install -e --id Oracle.JavaRuntimeEnvironment -v 8.0.3010.9 --accept-package-agreements
        winget install -e --id 7zip.7zip --accept-package-agreements
        winget install -e --id CodecGuide.K-LiteCodecPack.Full --accept-package-agreements
        winget install -e --id Microsoft.VCRedist.2015+.x86 --accept-package-agreements
        winget install -e --id Microsoft.VCRedist.2015+.x64 --accept-package-agreements
        winget install -e --id Microsoft.DirectX --accept-package-agreements
        winget install 9NKSQGP7F2NH --accept-package-agreements #Whatsapp
        winget install XP8BT8DW290MPQ --accept-package-agreements #Teams
        #winget install Microsoft.PowerBI --accept-package-agreements #PowerBI
        winget install Adobe.Acrobat.Reader.64-bit --accept-package-agreements #Adober Reader
        if($isAnydesk -eq "S" -or $isAnydesk -eq "s" -or $isAnydesk -eq "Y" -or $isAnydesk -eq "y"){ winget install -e --id AnyDeskSoftwareGmbH.AnyDesk --accept-package-agreements}
        #winget install 9WZDNCRFJ364 --accept-package-agreements #Skype
        #winget install XPFCG5NRKXQPKT --accept-package-agreements
        #winget install Avanquestpdfforge.PDFCreator-Free --accept-package-agreements --ignore-security-hash
    }
}

        #------------------------------Office
function RunOffice($isOffice, $isOfficeOp){
    if($isOffice -eq "S" -or $isOffice -eq "s" -or $isOffice -eq "Y" -or $isOffice -eq "y"){
        Write-Output "Instalando o Office"
        switch ($isOfficeOp) {
            1{Start-Process -FilePath "\\$servidor\TI\Softwares\Microsoft\Office\Desinstalar\setup.exe" -ArgumentList '/configure "\\$servidor\TI\Softwares\Microsoft\Office\Desinstalar\configuration.xml"' -Wait}
            2{Start-Process -FilePath "\\$servidor\softwares$\Microsoft\Office\Office2013\Office\setup64" -Wait}
            3{Start-Process -FilePath "\\$servidor\softwares$\Microsoft\Office\Office2019HB\Office\setup64.exe" -Wait}
            4{Start-Process -FilePath "\\$servidor\softwares$\Microsoft\Office\Office2021\Office\setup64.exe" -Wait}
        }
        Get-AppxPackage -AllUsers | Where-Object {$_.Name -Like '*OutlookForWindows*'} | Remove-AppxPackage -AllUsers -ErrorAction Continue
    }
}

        #------------------------------Tekla
function RunTekla($isTekla22, $isTekla24, $isTekla25){
    #------------------------------Tekla 2022
    if($isTekla22 -eq "S" -or $isTekla22 -eq "s" -or $isTekla22 -eq "Y" -or $isTekla22 -eq "y"){
        Write-Output "Instalando o Tekla 2022"
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Tekla 2022\TeklaStructures2022ServicePack12.exe" -Wait -ArgumentList "/s /v`"/qn INSTALLDIR=C:\TeklaStructures\ TSMODELDIR=C:\TeklaStructuresModels\ /lvoicewarmupx C:\Windows\Temp\TS2022_logfile.log`""
        Write-Output "Atualizando ambiente Default para a nova versao..."
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Tekla 2022\Env_Default_2022_120.exe" -Wait -ArgumentList "/s /v`"/qn /lvoicewarmupx C:\Windows\Temp\TS2022Default_logfile.log RUNATTSOPENING=true`""
        Write-Output "Atualizando ambiente Brazil para a nova versao..."
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Tekla 2022\Env_Brazil_2022_120.exe" -Wait -ArgumentList "/s /v`"/qn /lvoicewarmupx C:\Windows\Temp\TS2022Brazil_logfile.log RUNATTSOPENING=true`""
    }
    
    #------------------------------Tekla 2024
    if($isTekla24 -eq "S" -or $isTekla24 -eq "s" -or $isTekla24 -eq "Y" -or $isTekla24 -eq "y"){
        Write-Output "Instalando o Tekla 2024"
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Tekla 2024\TeklaStructures2024ServicePack4.1.exe" -Wait -ArgumentList "/s /v`"/qn INSTALLDIR=C:\TeklaStructures\ TSMODELDIR=C:\TeklaStructuresModels\ /lvoicewarmupx C:\Windows\Temp\TS2024_logfile.log`""
        Write-Output "Instalando ambiente Default para a nova versao..."
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Tekla 2024\Env_Default_2024_40.exe" -Wait -ArgumentList "/s /v`"/qn /lvoicewarmupx C:\Windows\Temp\TS2024Default_logfile.log RUNATTSOPENING=true`""
        Write-Output "Instalando ambiente Brazil para a nova versao..."
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Tekla 2024\Env_Brazil_2024_40.exe" -Wait -ArgumentList "/s /v`"/qn /lvoicewarmupx C:\Windows\Temp\TS2024Brazil_logfile.log RUNATTSOPENING=true`""
    }
        #------------------------------Tekla 2025
    if($isTekla25 -eq "S" -or $isTekla25 -eq "s" -or $isTekla25 -eq "Y" -or $isTekla25 -eq "y"){
        Write-Output "Instalando o Tekla 2025"
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Tekla 2025\TeklaStructures2025.exe" -Wait -ArgumentList "/s /v`"/qn INSTALLDIR=C:\TeklaStructures\ TSMODELDIR=C:\TeklaStructuresModels\ /lvoicewarmupx C:\Windows\Temp\TS2025_logfile.log`""
        Write-Output "Instalando ambiente Default para a nova versao..."
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Tekla 2025\Env_Default_2025.exe" -Wait -ArgumentList "/s /v`"/qn /lvoicewarmupx C:\Windows\Temp\TS2025Default_logfile.log RUNATTSOPENING=true`""
        Write-Output "Instalando ambiente Brazil para a nova versao..."
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Tekla 2025\Env_Brazil_2025.exe" -Wait -ArgumentList "/s /v`"/qn /lvoicewarmupx C:\Windows\Temp\TS2025Brazil_logfile.log RUNATTSOPENING=true`""
    }
}

        #------------------------------Trimble Connect
function RunTrimble($isTrimble){
    if($isTrimble -eq "S" -or $isTrimble -eq "s" -or $isTrimble -eq "Y" -or $isTrimble -eq "y"){
        Write-Output "Instalando o Trimble"
        Start-Process -FilePath "\\$servidor\Softwares$\Tekla\Trimble Connect\TrimbleConnectSetup-1.23.1.553-x64.exe" -Wait -ArgumentList "/s /v`"/qn`""
    }
}

        #------------------------------Autocad 2018
function RunAutocad($isAutocad18){
    if($isAutocad18 -eq "S" -or $isAutocad18 -eq "s" -or $isAutocad18 -eq "Y" -or $isAutocad18 -eq "y"){
        Write-Output "Instalando o Autocad 2018"
        Start-Process -FilePath "\\$servidor\softwares$\Autodesk\AutoCAD_2018_English_Win_64bit_r1_dlm_001_002.sfx.exe" -ArgumentList '-q' -Wait
    }
}

        #------------------------------Navisworks Freedom
function RunNavis($isNavis){
    if($isNavis -eq "S" -or $isNavis -eq "s" -or $isNavis -eq "Y" -or $isNavis -eq "y"){
        Write-Output "Navisworks Freedom"
        Start-Process -FilePath "\\$servidor\softwares$\Autodesk\Navisworks Fredom 2025\Autodesk_Navisworks_Freedom_2025_Win_64bit_db_001_002.exe" -Wait
    }
}

        #------------------------------Mfiles
function RunMfiles($isMfiles){
    if($isMfiles  -eq "S" -or $isMfiles -eq "s" -or $isMfiles -eq "Y" -or $isMfiles -eq "y"){
        Write-Output "Instalando Mfiles"
        Start-Process -FilePath "\\$servidor\Softwares$\M-Files\M-Files_Online_x64_ptb_client_24_8_13981_4_EV.msi" -Wait
    }
}

        #------------------------------True View
function RunTruview($isTrueview){
    if($isTrueview -eq "S" -or $isTrueview -eq "s" -or $isTrueview -eq "Y" -or $isTrueview -eq "y"){
        Write-Output "Instalando o True View"
        Start-Process -FilePath "\\$servidor\softwares$\Autodesk\Visualizador AutoCAD\TrueView2025\Setup.exe" --silent
    }
}

        #------------------------------Update do Windows
function RunUpdate($isUpdate){
    if($isUpdate -eq "S" -or $isUpdate -eq "s" -or $isUpdate -eq "Y" -or $isUpdate -eq "y") {
        Write-Host "Verificando atualizacoes do Windows..." -ForegroundColor Cyan

        # Verifica se o serviço Windows Update esta ativo
        $windowsUpdateService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
        if ($null -eq $windowsUpdateService -or $windowsUpdateService.Status -ne 'Running') {
            Write-Host "O serviço Windows Update esta desativado. Iniciando serviço..." -ForegroundColor Yellow
            Try {
                Start-Service -Name wuauserv -ErrorAction Stop
                Write-Host "Serviço Windows Update iniciado com sucesso" -ForegroundColor Green
            } Catch {
                Write-Host "Erro ao iniciar o serviço Windows Update. Tente ativa-lo manualmente" -ForegroundColor Red
                Pause
                return
            }
        }

        # Verifica se o módulo PSWindowsUpdate esta instalado
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Host "O módulo PSWindowsUpdate nao esta instalado. Instalando..." -ForegroundColor Yellow
            Try {
                Install-Module PSWindowsUpdate -Force -Confirm:$false -ErrorAction Stop
            } Catch {
                Write-Host "Erro ao instalar o módulo PSWindowsUpdate. Verifique sua conexao com a internet" -ForegroundColor Red
                Pause
                return
            }
        }

        # Importa o módulo PSWindowsUpdate
        Try {
            Import-Module PSWindowsUpdate -ErrorAction Stop
        } Catch {
            Write-Host "Erro ao importar o módulo PSWindowsUpdate. Verifique se ele esta instalado corretamente" -ForegroundColor Red
            Pause
            return
        }

        # Obtém a lista de atualizacoes disponíveis
        Try {
            $atualizacoes = Get-WindowsUpdate -ErrorAction Stop
        } Catch {
            Write-Host "Erro ao buscar atualizacoes. Certifique-se de que o Windows Update esta ativado" -ForegroundColor Red
            Pause
            return
        }

        if ($null -eq $atualizacoes -or $atualizacoes.Count -eq 0) {
            Write-Host "O Windows ja esta atualizado Nenhuma atualizacao pendente." -ForegroundColor Green
        } else {
            Write-Host "As seguintes atualizacoes estao disponíveis:" -ForegroundColor Yellow
            $atualizacoes | ForEach-Object { Write-Host "- $($_.Title)" -ForegroundColor Cyan }

            Write-Host "`nIniciando a instalacao das atualizacoes..." -ForegroundColor Yellow
            Try {
                Install-WindowsUpdate -AcceptAll -AutoReboot -ErrorAction Stop
                Write-Host "Atualizacoes instaladas com sucesso" -ForegroundColor Green
            } Catch {
                Write-Host "Erro ao instalar atualizacoes. Tente novamente mais tarde" -ForegroundColor Red
            }
        }
    }
}

        #------------------------------Otimizar o Windows
function RunCleanup($isCleandisk){
    if($isCleandisk -eq "S" -or $isCleandisk -eq "s" -or $isCleandisk -eq "Y" -or $isCleandisk -eq "y"){
        Write-Host "Otimizando o Windows..." -ForegroundColor Cyan
        Write-Host "Executando a limpeza de disco..." -ForegroundColor Yellow
            try {
                Start-Process -FilePath "cleanmgr.exe" -ArgumentList "sagerun:1" -NoNewWindow -Wait
                Write-Host "Limpeza de disco concluida" -ForegroundColor Green
            } Catch {
                Write-Host "Erro ao executar a limpeza de disco: $_" -ForegroundColor Red
            }
            Write-Host "Realizando a exclusao dos arquivos da pasta temp" -ForegroundColor Yellow
            try {
                Remove-Item -Path "$env:TEMP\*" -Force -Recurse -ErrorAction Stop
                Write-Host "Arquivos temporarios removidos com sucesso" -ForegroundColor Green
            } Catch {
                Write-Host "Erro ao tentar excluir os arquivos da pasta TEMP: $_" -ForegroundColor Red
            }
            Write-Host "Realizando a exclusao dos arquivos das lixeiras" -ForegroundColor Yellow
            try {
                Clear-RecycleBin -DriveLetter C -Force
                Write-Host "Arquivos da Lixeira removidos com sucesso" -ForegroundColor Green
            } Catch {
                Write-Host "Erro ao tentar excluir os arquivos da pasta LIXEIRA: $_" -ForegroundColor Red
            }
    }
}

function RunPassGenerator {
    # Requisitos
    $minLength = 9
    $nchar = @('1','i','I','o','0','l','L','|','^','~','ç','_')
    $nconter = @('metasa', 'Metasa', 'Marau', 'marau')
    $lowercase = ([char[]]'abcdefghjkmnpqrstuvwxyz')
    $uppercase = ([char[]]'ABCDEFGHJKMNPQRSTUVWXYZ')
    $digits    = ([char[]]'23456789')
    $specials  = ([char[]]'@#$%&*+=!?-')
    $allChars = $lowercase + $uppercase + $digits + $specials

    do {
        $categories = @()
        $categories += ,($lowercase | Get-Random)
        $categories += ,($uppercase | Get-Random)
        $categories += ,($digits | Get-Random)
        if((Get-Random -Minimum 0 -Maximum 2) -eq 1) {$categories += ,($specials | Get-Random)}

        while ($categories.Count -lt $minLength) {
            $char = $allChars | Get-Random
            if (-not ($nchar -contains $char)) {
                $categories += $char
            }
        }

        $password = ($categories | Sort-Object {Get-Random}) -join ''

        $valid = $true
        foreach ($word in $nconter) {
            if ($password.ToLower().Contains($word.ToLower())) {
                $valid = $false
                break
            }
        }

        if (-not $valid) {
            Write-Host "Senha inválida gerada, tentando novamente..."
        } else {
            Write-Host "-------Senha Gerada--------"
            Write-Host "        $password" -ForegroundColor Green
            Write-Host "---------------------------"
            $opend = Read-Host "Pressione ENTER para voltar ou G para gerar novamente: "
            if ($opend -eq "G" -or $opend -eq "g") {
                $valid = $false
            } else {
                Write-Host "Pressione qualquer tecla para continuar..."
                break
            }
        }
    } while (-not $valid)
}

function CreatePassword{
    $usuario = Read-Host "        Informe o usuário para gerar a senha: "
    $senha = "M" + $usuario[0] + $usuario[-1] + $usuario[1] + $usuario.Length + ($usuario.Length *2)
    Write-Host "-------Senha Gerada--------"
    Write-Host "        $senha" -ForegroundColor Green
    Write-Host "---------------------------"
}

function getInfoMachine() {

    function Format-WMIDate {
        param (
            [string]$wmiDate
        )
        $date = [System.Management.ManagementDateTimeConverter]::ToDateTime($wmiDate)
        return $date.ToString("yyyy-MM-dd HH:mm:ss")
    }

    # Coleta de informações do computador
    $computerInfo = @{
        "Nome do Computador" = hostname
        "Endereço IP" = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -ne "Loopback Pseudo-Interface 1" }).IPAddress
        "Endereço MAC" = (Get-NetAdapter | Where-Object { $_.Status -eq "Up" }).MacAddress
        "Velocidade de Rede" = (Get-NetAdapter | Where-Object { $_.Status -eq "Up" }).LinkSpeed
        "Uso de CPU (%)" = (Get-WmiObject -Class Win32_Processor).LoadPercentage
        "Uso de RAM (%)" = [math]::Round(((Get-WmiObject -Class Win32_OperatingSystem).TotalVisibleMemorySize - (Get-WmiObject -Class Win32_OperatingSystem).FreePhysicalMemory) / (Get-WmiObject -Class Win32_OperatingSystem).TotalVisibleMemorySize * 100, 2)
        "Uso de Disco (%)" = [math]::Round(((Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3").Size - (Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3").FreeSpace) / (Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3").Size * 100, 2)
        "Versão do BIOS" = (Get-WmiObject -Class Win32_BIOS).SMBIOSBIOSVersion
        "Inicialização do Sistema" = Format-WMIDate ((Get-WmiObject -Class Win32_OperatingSystem).LastBootUpTime)
        "Data da Instalação Original" = Format-WMIDate ((Get-WmiObject -Class Win32_OperatingSystem).InstallDate)
        "Sistema Operacional" = (Get-WmiObject -Class Win32_OperatingSystem).Caption
        "Versão do Sistema Operacional" = (Get-WmiObject -Class Win32_OperatingSystem).Version
        "Fabricante do Computador" = (Get-WmiObject -Class Win32_ComputerSystem).Manufacturer
        "Modelo do Computador" = (Get-WmiObject -Class Win32_ComputerSystem).Model
        "Serial Number" = Get-WmiObject -Class Win32_Bios | select SerialNumber
    }

    $userProfiles = Get-WmiObject -Class Win32_UserProfile
    $userLogonInfo = @()
    foreach ($profile in $userProfiles) {
        $userName = $profile.LocalPath.Split('\')[-1]
        try {
            $lastAccessTime = Format-WMIDate $profile.LastUseTime
        }
        catch {
            $lastAccessTime = "Não foi possível obter a data do último acesso"
        }
        $logonInfo = [PSCustomObject]@{
            UserName = $userName
            UserProfilePath = $profile.LocalPath
            LastAccessTime = $lastAccessTime
        }
        $userLogonInfo += $logonInfo
    }

    $computerInfo
    $userLogonInfo | Format-Table -AutoSize
    Write-Host "Pressione qualquer tecla para continuar..."
    [void][System.Console]::ReadKey($true)
}

function runBackup(){ 
    # Solicitar ao usuário se deseja copiar somente um usuário 
    $op = Read-Host "Para copiar somente um usuário informe: fulano.fulano, caso contrário digite: N" 
    $origem = Read-Host "Digite o caminho da máquina na rede (ex: \\maquina-rede\C$\Users)" 
    $destino = Read-Host "Digite o caminho para salvar o backup (ex: C:\BackupUsuarios)" 
    $opext = Read-Host "Deseja ignorar arquivos com extensão (.exe, .tmp, .bak, .log) S/N ? " 
    
    # Validar caminhos 
    if (!(Test-Path -Path $origem)) { 
        Write-Host "Erro: O caminho de origem fornecido não é válido. Verifique e tente novamente." 
        return 
    } 

    if (!(Test-Path -Path $destino)) { 
        Write-Host "Criando diretório de destino..." 
        New-Item -Path $destino -ItemType Directory | Out-Null 
    } 

    if($opext -eq "S" -or $opext -eq "s" -or $opext -eq "Y" -or $opext -eq "y"){ 
        $extensoesIgnoradas = @("*.exe", "*.tmp", "*.bak", "*.log") 
        $xfParam = "/XF " + ($extensoesIgnoradas -join " ")
    } else { 
        $xfParam = ""
    } 

    # Loop para copiar pastas Desktop e Documents de cada usuário 
    Get-ChildItem -Path $origem -Directory | ForEach-Object { 
        $usuario = $_.Name 

        if ($op -eq "N" -or $op -eq "n" -or $op -eq $usuario) { 
            Write-Host "Copiando Área de Trabalho e Documentos do usuário $usuario..." 

            $pastas = @("Desktop", "Documents") 
            foreach ($pasta in $pastas) { 
                $origemPasta = Join-Path -Path $_.FullName -ChildPath $pasta 
                $destinoPasta = Join-Path -Path $destino -ChildPath "$usuario\$pasta" 

                if (Test-Path -Path $origemPasta) { 
                    $result = robocopy $origemPasta $destinoPasta /E $xfParam
                    if ($LASTEXITCODE -ge 8) { 
                        Write-Host "Erro ao copiar $pasta do usuário $usuario." 
                    } 
                } 
            } 
        } 
    } 

    Write-Host "Backup concluído!" 
}

# Loop principal para o menu
do {
    $op = printMenu
    $opStandard = $null
    $opEng = $null
    $opRemote = $null
    $opConfig = $null

    switch ($op) {
        1 {
            $opStandard = readStandard
            $opEng = readEng
            $opRemote = readRemote
            $opConfig = readConfig
            Write-Host " . . . Configuracoes escolhidas para Instalacao Inicial:" -ForegroundColor Green
        }
        2 {
            $opStandard = readStandard
            Write-Host " . . . Configuracoes escolhidas para Computadores padrao:" -ForegroundColor Green
        }
        3 {
            $opEng = readEng
            Write-Host "`n . . . Configuracoes escolhidas para Computadores da Engenharia:" -ForegroundColor Green
        }
        4 {
            $opRemote = readRemote
            Write-Host "`n . . . Configuracoes escolhidas para Computadores Remotos:" -ForegroundColor Green
        }
        5 {
            $opConfig = readConfig
            Write-Host "`n . . . Configuracoes escolhidas para Otimizacoes:" -ForegroundColor Green
        }
        6 {
            
            Write-Host "`n . . . Selecione qual executar" -ForegroundColor Green
            readSelect
        }
        7 {
            Write-Host "`n . . . Gerando Senha..." -ForegroundColor Green
            #RunPassGenerator
            CreatePassword
        }
        8 {
            Write-Host "`n . . . Buscando dados..." -ForegroundColor Green
            getInfoMachine
        }
        9 {
            Write-Host "`n . . . Executar Backup" -ForegroundColor Green
            runBackup
        }
        10 {
            Write-Host "`n . . . Saindo do programa..." -ForegroundColor Green
        }
        Default {
            Write-Host "`n . . . Opcao invalida Por favor, tente novamente." -ForegroundColor Red
        }
    }

    RunWinget $opStandard.Winget
    RunSapiens $opStandard.Sapiens
    RunVetorH $opStandard.VetorH
    RunTotvs $opStandard.Datasul $opStandard.RM
    RunPrograms $opStandard.SoftwaresBasicos $opRemote.Anydesk
    RunTrauma $opStandard.TraumaZero
    RunOracle $opStandard.Oracle
    RunSymantec $opStandard.Antivirus
    RunOffice $opStandard.Office $opStandard.OfficeVersao
    RunEnergy $opConfig.Energy
    RunUpdate $opConfig.Update
    RunCleanup $opConfig.Cleandisk
    RunRdp $opRemote.Rdp
    RunTrimble $opEng.Trimble
    RunTekla $opEng.Tekla22 $opEng.Tekla24 $opEng.Tekla25
    RunAutocad $opEng.Autocad18
    RunNavis $opEng.Navis
    RunMfiles $opEng.Mfiles
    RunTruview $opEng.Trueview

    Write-Host "Pressione qualquer tecla para continuar..."
    [void][System.Console]::ReadKey($true)

} while ($op -ne "0")


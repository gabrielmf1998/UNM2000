$asciiArt = @'

  _    _ _   _ __  __ ___   ___   ___   ___  
 | |  | | \ | |  \/  |__ \ / _ \ / _ \ / _ \ 
 | |  | |  \| | \  / |  ) | | | | | | | | | |
 | |  | | . ` | |\/| | / /| | | | | | | | | |
 | |__| | |\  | |  | |/ /_| |_| | |_| | |_| |
  \____/|_| \_|_|  |_|____|\___/ \___/ \___/ 
                                             
                                             

'@
Write-Host $asciiArt -ForegroundColor Blue

#Feito Por Gabriel Marques Ferrarezi
Write-Host "Iniciando..." -ForegroundColor Green

#Vai pedir senha para o usuário

Write-Host "Digite a senha para liberação: " -ForegroundColor Yellow -NoNewline
$securePassword = Read-Host -AsSecureString
$unsecurePassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
)
if ($unsecurePassword -eq "PRIMEIRA_SENHA") {
} else {
    Write-Host "Senha incorreta!" -ForegroundColor Red
	Start-Sleep 1
	exit
}
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 1


Write-Host "Validando disco para instalação..." -ForegroundColor Green
Start-Sleep 2

#Vai checar se o disco é compatível com UNM2000
Add-Type -AssemblyName System.Windows.Forms
$drive = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "D:" }

if (-not $drive) {
    [System.Windows.Forms.MessageBox]::Show(
        "Unidade D: não existe!`nVerifique e inicie o programa novamente!",
        "Erro de Instalação",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit
}

switch ($drive.DriveType) {
    2 {
        [System.Windows.Forms.MessageBox]::Show(
            "Unidade D: é uma unidade de Disquete.`nNão é válido para instalação do UNM2000!",
            "Erro de Instalação",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null        
        exit
    }
    5 {
        [System.Windows.Forms.MessageBox]::Show(
            "Unidade D: é um CD/DVD.`nNão é válido para instalação do UNM2000!",
            "Erro de Instalação",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        exit
    }
    3 {
        if ($drive.FileSystem -ne "NTFS") {
            [System.Windows.Forms.MessageBox]::Show(
                "Unidade D: encontrada, porém não é um HDD/SSD.`nNão é válido para instalação do UNM2000!",
                "Erro de Instalação",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            exit
        }
    }
    default {
        [System.Windows.Forms.MessageBox]::Show(
            "Unidade D: tipo de unidade desconhecido.`nNão é válido para instalação do UNM2000!",
            "Erro de Instalação",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        exit
    }
}

$freeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
if ( $freeGB -le 550) {
	$result = [System.Windows.Forms.MessageBox]::Show(
    "Unidade D: encontrada!`nMas possui apenas $freeGB GB de espaço livre.`nRecomendado para UNM2000 é 600 GB `n`nDeseja continuar com a instalação? ",
    "Confirmação",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Warning
)
if ($result -eq [System.Windows.Forms.DialogResult]::No) {
	Write-Host "Fechando Programa..." -ForegroundColor Red
	Start-Sleep 1
    exit
}
}

Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 2

#Vai solicitar para o usuario confirmar o IP da máquina

Add-Type -AssemblyName System.Windows.Forms
Write-Host "Verificando informações de Rede..." -ForegroundColor Green
ncpa.cpl
$resposta = [System.Windows.Forms.MessageBox]::Show(
    "Confirme se a máquina possui IP-Fixo!`nÉ importante para funcionamento do UNM2000 fixar endereço.`n`nTem certeza que o IP está Fixo?",
    "Confirmação",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)
if ($resposta -eq [System.Windows.Forms.DialogResult]::Yes) {
    $resposta2 = [System.Windows.Forms.MessageBox]::Show(
    "Uma vez definido IP na máquina e instalado UNM, o mesmo não pode mais ser alterado!`n`nTem ABSOLUTA certeza que o endereço da Máquina/VM está Fixo?",
    "Confirmação",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Warning
	)
} else {
    Write-Host "Encerrando programa..." -ForegroundColor Red
	Start-Sleep 2
	exit
}
    if ($resposta2 -eq [System.Windows.Forms.DialogResult]::No) {
	Write-Host "Encerrando programa..." -ForegroundColor Red
	Start-Sleep 2
	exit
}
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 3


#Vai fazer download do 7zip e instalar
Write-Host "Baixando e instalando 7Zip..." -ForegroundColor Green
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
$ScriptPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$ScriptDir = Split-Path -Parent $ScriptPath
$installerPath = "$env:TEMP\7zip.msi"
Invoke-WebRequest -Uri "LINK_DIRETO_7ZIP" -OutFile $installerPath
Start-Process msiexec.exe -ArgumentList "/i `"$installerPath`" /quiet /norestart" -Wait
Get-ChildItem "C:\Program Files\7-Zip\7z.exe"
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 2

#Vai pedir senha para o usuário
Write-Host "Digite a senha para download do UNM2000: " -ForegroundColor Yellow -NoNewline
$securePassword = Read-Host -AsSecureString
$unsecurePassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
)
if ($unsecurePassword -eq "SEGUNDA_SENHA") {
} else {
    Write-Host "Senha incorreta!" -ForegroundColor Red
	Start-Sleep 1
	exit
}
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 1


#Vai baixar o UNM2000
Write-Host "Baixando UNM2000..." -ForegroundColor Green
Write-Host "Aguarde, o download pode demorar alguns minutos..." -ForegroundColor Yellow
$ScriptPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$ScriptDir = Split-Path -Parent $ScriptPath
$installerPath = "UNM2000_Full.zip"
Invoke-WebRequest -Uri "link_direto_instalacao_unm2000" -OutFile $installerPath
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 2

#Vai checar se o arquivo existe
Write-Host "Validando arquivo UNM2000_Full.zip..." -ForegroundColor Green
$arquivo = "UNM2000_Full.zip"
$tamanhoEsperado = 2892222218
if (-not (Test-Path $arquivo)) {
    [System.Windows.Forms.MessageBox]::Show(
            "Arquivo 'UNM2000_Full.zip' não encontrado!`nMáquina possui internet?",
            "Erro de Instalação",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        Write-Host "Encerrando programa..." -ForegroundColor Red
	    Start-Sleep 2
	    exit
}
$infoArquivo = Get-Item $arquivo
if ($infoArquivo.Length -ne $tamanhoEsperado) {
    [System.Windows.Forms.MessageBox]::Show(
            "Arquivo encontrado, mas está incorreto ou corrompido.`nExecute o script novamente.",
            "Erro de Instalação",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        Write-Host "Encerrando programa..." -ForegroundColor Red
	    Start-Sleep 2
	    exit
}
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 3


#Vai extrair o UNM2000
Write-Host "Extraindo UNM2000_Full.zip..." -ForegroundColor Green
$ScriptPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$ScriptDir = Split-Path -Parent $ScriptPath
$sevenZip = "C:\Program Files\7-Zip\7z.exe"
$zip = Join-Path $ScriptDir "UNM2000_Full.zip"
& "$sevenZip" x "$zip" -o"$ScriptDir" -y
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 2

#Vai instalar o Winrar
Write-Host "Instalando Winrar..." -ForegroundColor Green
$ScriptPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$ScriptDir = [System.IO.Path]::GetDirectoryName($ScriptPath)
$instaladorwinrar = Join-Path $ScriptDir ".\UNM 2000 - Files\UNM2000\winrar-x64-700br.exe"
Start-Process -FilePath $instaladorwinrar -ArgumentList "/S" -Wait
Get-ChildItem "C:\Program Files\WinRAR\WinRAR.exe"
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 2

#Vai instalar Java JDK
Write-Host "Instalando JDK..." -ForegroundColor Green
$ScriptPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$ScriptDir = [System.IO.Path]::GetDirectoryName($ScriptPath)
$installerjdk = Join-Path $ScriptDir ".\UNM 2000 - Files\UNM2000\jdk-8u201-windows-x64.exe"
Start-Process -FilePath $installerjdk -ArgumentList "/s" -Verb RunAs -Wait
Get-ChildItem "C:\Program Files\Java"
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 2

#Vai importar os arquivos no firewall
Write-Host "Criando regras no Firewall..." -ForegroundColor Green
$firewallConfigPath = ".\UNM 2000 - Files\UNM2000\unm.wfw"
netsh advfirewall import "$firewallConfigPath"
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 3

#Vai extrair o MySQL e criar ambos UNM e Pas
Write-Host "Extraindo arquivos MySQL..." -ForegroundColor Green
$zipPath = ".\UNM 2000 - Files\UNM2000\MySQL5.7.35_x64.rar"
$winRAR = "C:\Program Files\WinRAR\Rar.exe"
$destUnm = "D:\"

& "$winRAR" x -o+ "$zipPath" "$destUnm\"

$origem = "D:\MySQL"
$novoNome = "MySQL_Unm"
Rename-Item -Path $origem -NewName $novoNome

& "$winRAR" x -o+ "$zipPath" "$destUnm\"
$origem = "D:\MySQL"
$novoNome = "MySQL_Pas"
Rename-Item -Path $origem -NewName $novoNome
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 3

#Vai instalar o MySQL_UNM e MySQL_Pas
$MySQL_Unm = "D:\MySQL_Unm\install_mysql.bat"
$MySQL_Pas = "D:\MySQL_Pas\install_mysql.bat"
Write-Host "Instalando MySQL_Unm, Lembre-se de DIGITAR 1" -ForegroundColor Yellow
#& $MySQL_Unm
Start-Process -FilePath $MySQL_Unm -ArgumentList "/s" -Verb RunAs -Wait
Start-Sleep 2
Write-Host "Instalando MySQL_Pas, Lembre-se de DIGITAR 0" -ForegroundColor Yellow
#& $MySQL_Pas
Start-Process -FilePath $MySQL_Pas -ArgumentList "/s" -Verb RunAs -Wait
Start-Sleep 2
Write-Host "[OK]" -ForegroundColor Green

Start-Sleep 2
#Vai adicionar as variáveis de ambiente
Write-Host "Adicionando variáveis de ambiente no Windows..." -ForegroundColor Green
$pathsToAdd = @(
    "%JAVA_HOME%\bin",
    "D:\MySQL_Unm\bin",
    "D:\MySQL_Pas\bin"
)

$existingPath = [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::Machine)

foreach ($path in $pathsToAdd) {
    if (-not ($existingPath -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ -eq $path })) {
        $existingPath += ";$path"
    }
}

[Environment]::SetEnvironmentVariable("Path", $existingPath, [EnvironmentVariableTarget]::Machine)
[Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::Machine)
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 3

#Vai instalar o UNM2000
$ScriptPath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$ScriptDir = [System.IO.Path]::GetDirectoryName($ScriptPath)
$installerUNM = Join-Path $ScriptDir ".\UNM 2000 - Files\UNM2000\unm2000-1.0-windows-installer-20220930-157490.exe"
Write-Host "Instalando UNM2000.exe..." -ForegroundColor Green
Write-Host "Aguarde a instalação concluir do UNM2000!" -ForegroundColor Yellow
Start-Process -FilePath $installerUNM -Verb RunAs -Wait
Write-Host "[OK]" -ForegroundColor Green

#Vai criar database
Write-Host "[OK]" -ForegroundColor Green
Write-Host "Execute o script Createdatabase.bat manualmente..." -ForegroundColor Yellow
Write-Host "Não esqueca que deve ser em modo ADM!" -ForegroundColor Yellow
explorer "D:\unm2000\script\"
do {
    do {
        Write-Host "Executou o script Createdatabase.bat? s/n" -ForegroundColor Yellow
        Write-Host "Confirmação: " -ForegroundColor Yellow -Nonewline
        $resposta = Read-Host
        $resposta = $resposta.Trim().ToUpper()

        if ($resposta -ne "s" -and $resposta -ne "n") {
            Write-Host "Resposta inválida!" -ForegroundColor Red
        }
    } while ($resposta -ne "s" -and $resposta -ne "n")

    if ($resposta -eq "s") {
        Write-Host "[OK]" -ForegroundColor Green
        break  # Sai do loop principal
    } elseif ($resposta -eq "n") {
        explorer "D:\unm2000\script\"
        Write-Host "Confirme e execute Createdatabase.bat novamente!" -ForegroundColor Red
        Start-Sleep 2
    }
} while ($true)

#Vai Subir serviços
Write-Host "Subindo serviços...." -ForegroundColor Green
$StartServicesbat = "D:\unm2000\script\startAllService.bat"
Start-Process -FilePath $StartServicesbat -Verb RunAs -Wait
Write-Host "[OK]" -ForegroundColor Green

#Deleta arquivos restantes
Write-Host "Fazendo Limpeza..." -ForegroundColor Green
Start-Sleep 1
Remove-Item -Path ".\UNM 2000 - Files\" -ErrorAction SilentlyContinue -Recurse -Force -Confirm:$false
Remove-Item -Path ".\UNM2000_Full.zip" -ErrorAction SilentlyContinue
Write-Host "[OK]" -ForegroundColor Green
Start-Sleep 2

#Vai mostrar os logs para o usuario
Write-Host "Instalação Concluída!" -ForegroundColor Green
Write-Host "Acompanhe os logs WaitStart abaixo..." -ForegroundColor Green
$datadehoje = Get-Date -Format "yyyy-MM-dd"
$caminho = "D:\unm2000\server\log\unm_service_monitor.$datadehoje.log"
Get-Content -Path $caminho -Wait
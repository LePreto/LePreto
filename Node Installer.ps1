cls
#CHCP 65001
$baseUrl = "https://nodejs.org/"
$tempFilesPath = (Join-Path -Path (Get-Location) -Child "\tempFiles")
$userPath = ($env:USERPROFILE)
$extractDest
$startRequest
$links
$keyWord = "dist"
$currVer = 0
$lastVer = 0
$lastVerUrl  = ""
$fileRequest
$filesLinks
$fileName
$fileUrl
$extractShell = New-Object -ComObject Shell.Application 
$extractFiles
$extractFolderName
$extractDestFolder = ($userPath+"\bin\nodejs")
$ambientVar

# Get object methods:
# Write-Host $startRequest | Get-Member

Write-Host 'Baixando site nodejs.org'
$startRequest = Invoke-WebRequest -URI $baseUrl -UseBasicParsing
$links = $startRequest.Links

Write-Host 'buscando links...'
foreach ($link in $links) {
	if($link.href.ToLower().IndexOf($keyWord) -eq -1 -Or $link.href.ToLower().IndexOf("api") -gt -1){Continue}
	
	$lastVer = ($link.href.Substring($link.href.ToLower().IndexOf($keyWord)+$keyWord.Length+2, 2)/1)
	if($lastVer -gt $currVer){
		$currVer = $lastVer
		$lastVerUrl = $link.href
	}
}
# Write-Host $currVer
# Write-Host $lastVerUrl

Write-Host "Link encontrado: $lastVerUrl"

Write-Host 'Navegando para a página de arquivos...'
$fileRequest = Invoke-WebRequest -URI $lastVerUrl -UseBasicParsing
$filesLinks = $fileRequest.links

Write-Host 'buscando links de instalações...'
# Obtém o link a ser baixado
foreach ($fileLink in $filesLinks) {
	if($fileLink.href.ToLower().IndexOf("zip") -eq -1 -Or $fileLink.href.ToLower().IndexOf("64") -eq -1){Continue}
	$fileName = $fileLink.href
	$fileUrl = $($lastVerUrl + $fileLink.href)
}

# Testa se o "tempFilesPath" existe, caso contrário cria-o
if (-not(Test-Path -Path $tempFilesPath)) {
	New-Item -Path $tempFilesPath -ItemType Directory
}

# Testa se o arquivo já foi baixado, caso contrário baixa-o
if (-not(Test-Path -Path ($tempFilesPath +"\"+ $fileName))) {
	Write-Host 'Baixando instalação...'
	Invoke-WebRequest -Uri $fileUrl -OutFile ($tempFilesPath +"\"+ $fileName) -UseBasicParsing
}

$extractFiles = $extractShell.Namespace($tempFilesPath +"\"+ $fileName).Items() 
$extractFolderName = $extractFiles.Item(0).Path.Substring($extractFiles.Item(0).Path.lastIndexOf("\")+1)

# Caso não exista o caminho pra extrair os arquivos, cria-se
if (-not(Test-Path -Path ($extractDestFolder))) {
	New-Item -Path $extractDestFolder -ItemType Directory
}

# Caso exista o caminho da pasta que contenha os dados extraídos, apaga-se
if(Test-Path -Path ($extractDestFolder+"\"+$extractFolderName)){
	Write-Host "Removendo arquivos existentes na pasta destido..."
	rm -r ($extractDestFolder+"\"+$extractFolderName)
}

# No caso de alterar o nome da pasta obtida dentro do arquivo ZIP baixado
# Rename-Item -Path ($extractDestFolder+"\"+$extractFolderName) -NewName "monday_file.txt"

Write-Host "Extraindo arquivos baixados..."
Write-Host ""
Expand-Archive -Path ($tempFilesPath +"\"+ $fileName) -DestinationPath ($extractDestFolder) -Force

# Variáveis para atualizar o 
$domain, $userName = (Get-WmiObject -Class Win32_ComputerSystem).UserName -split '\\', 2
$user = [System.Security.Principal.NTAccount]::new($domain, $userName)
$sid  = $user.Translate([System.Security.Principal.SecurityIdentifier]).Value
$ambientVar = (New-Object -ComObject WScript.Shell).RegRead("HKEY_USERS\$sid\Environment\Path")

# Atualiza a variável de ambiente NODEJS_HOME do usuário atual
Set-ItemProperty -Path HKCU:\Environment -Name NODEJS_HOME -Value ($extractDestFolder+"\"+$extractFolderName)

# Atualiza o template da variavel Path com %NODEJS_HOME% caso não exista
if($ambientVar.IndexOf('%NODEJS_HOME%') -eq -1){
	Set-ItemProperty -Path HKCU:\Environment -Name Path -Value ($ambientVar+";%NODEJS_HOME%")
}

Write-Host 'Removendo pasta "tempFiles", contendo arquivos baixados...'
rm -r $tempFilesPath


Write-Host 'Instalação concluída com sucesso!'

# Atualiza a variavel de ambiente
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","User")
pause

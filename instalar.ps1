$ErrorActionPreference = "Stop"
$Pasta = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Pasta

Write-Host "CONFIGURACAO DO BI DE OS" -ForegroundColor Cyan
if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw "Instale o Git for Windows: https://git-scm.com/download/win" }
if (-not (Get-Command py -ErrorAction SilentlyContinue) -and -not (Get-Command python -ErrorAction SilentlyContinue)) { throw "Instale o Python 3: https://www.python.org/downloads/" }

$Repo = Read-Host "Cole a URL HTTPS do repositorio vazio (ex.: https://github.com/usuario/bi-os-ero.git)"
if (-not (Test-Path ".git")) { git init -b main }
if (git remote get-url origin 2>$null) { git remote set-url origin $Repo } else { git remote add origin $Repo }

$Nome = git config --global user.name
if (-not $Nome) { git config --global user.name (Read-Host "Seu nome para o Git") }
$Email = git config --global user.email
if (-not $Email) { git config --global user.email (Read-Host "Seu e-mail do GitHub") }

if (Get-Command py -ErrorAction SilentlyContinue) {
  py -3 "$Pasta\atualizar_dados.py" --sem-push
} else {
  python "$Pasta\atualizar_dados.py" --sem-push
}
if ($LASTEXITCODE -ne 0) { throw "Não foi possível ler o Report.xlsx. Confira config.json e o acesso à rede." }
git add .
git commit -m "Publica BI de acompanhamento de OS"
git push -u origin main

$Config = Get-Content "$Pasta\config.json" -Raw | ConvertFrom-Json
$Minutos = [int]$Config.intervalo_minutos
$Acao = New-ScheduledTaskAction -Execute "$Pasta\sincronizar.bat" -WorkingDirectory $Pasta
$Inicio = (Get-Date).AddMinutes(2)
$Gatilho = New-ScheduledTaskTrigger -Once -At $Inicio -RepetitionInterval (New-TimeSpan -Minutes $Minutos)
$Usuario = "$env:USERDOMAIN\$env:USERNAME"
Register-ScheduledTask -TaskName "Atualizar BI de OS" -Action $Acao -Trigger $Gatilho -User $Usuario -RunLevel Limited -Force | Out-Null

Write-Host "`nConfiguracao concluida." -ForegroundColor Green
Write-Host "Agora ative o GitHub Pages em: Settings > Pages > Deploy from a branch > main > /(root)."
Write-Host "O site sera atualizado a cada $Minutos minutos enquanto este computador estiver ligado e conectado a rede."
Read-Host "Pressione ENTER para fechar"

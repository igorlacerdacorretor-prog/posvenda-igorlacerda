# Atualiza as páginas de pós-venda a partir da planilha e publica no GitHub Pages.
# Uso:  .\publicar.ps1
#       .\publicar.ps1 "C:\caminho\Outra Planilha.xlsx"   (para usar outro arquivo)

param(
    [string]$Planilha = "C:\Users\igorl\Desktop\Controle_de_Clientes_Igor_Lacerda_Imoveis_OTIMIZADO_3.xlsx"
)

$ErrorActionPreference = "Stop"

Write-Host "Gerando páginas a partir de '$Planilha'..." -ForegroundColor Cyan
python gerar_paginas.py $Planilha
if ($LASTEXITCODE -ne 0) {
    Write-Host "O script falhou. Corrija o erro acima antes de publicar." -ForegroundColor Red
    exit 1
}

git add docs/data "codigos_clientes.csv"

$mudancas = git status --porcelain -- docs/data codigos_clientes.csv
if (-not $mudancas) {
    Write-Host "Nenhuma mudança nova em relação ao que já está publicado. Nada a enviar." -ForegroundColor Yellow
    exit 0
}

Write-Host "Enviando atualização para o GitHub..." -ForegroundColor Cyan
git commit -m "Atualiza pos-venda - $(Get-Date -Format 'dd/MM/yyyy HH:mm')"
git push

Write-Host "Publicado! As páginas devem atualizar em ~1 minuto no GitHub Pages." -ForegroundColor Green

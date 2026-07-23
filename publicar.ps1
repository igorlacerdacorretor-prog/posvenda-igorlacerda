# Atualiza as páginas de pós-venda a partir da planilha e publica no GitHub Pages.
# Uso:  .\publicar.ps1 "C:\caminho\Pós Venda 2026.xlsx"

param(
    [Parameter(Mandatory = $true)]
    [string]$Planilha
)

$ErrorActionPreference = "Stop"

Write-Host "Gerando páginas a partir de '$Planilha'..." -ForegroundColor Cyan
python gerar_paginas.py $Planilha
if ($LASTEXITCODE -ne 0) {
    Write-Host "O script falhou. Corrija o erro acima antes de publicar." -ForegroundColor Red
    exit 1
}

Write-Host "Enviando atualização para o GitHub..." -ForegroundColor Cyan
git add docs/data "codigos_clientes.csv"
git commit -m "Atualiza pos-venda - $(Get-Date -Format 'dd/MM/yyyy HH:mm')"
git push

Write-Host "Publicado! As páginas devem atualizar em ~1 minuto no GitHub Pages." -ForegroundColor Green

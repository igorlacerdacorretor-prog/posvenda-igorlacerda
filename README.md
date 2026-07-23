# Pós-Venda | Igor Lacerda Imóveis

Página de acompanhamento de pós-venda que você envia por link (WhatsApp) para
cada cliente comprador. Cada cliente só vê o processo dele, através de um
código único e imprevisível na URL — não é possível adivinhar o link de
outro cliente.

## Como funciona

- `gerar_paginas.py` lê a aba **"Pós Venda 2026"** do seu Excel e gera um
  arquivo `docs/data/<codigo>.json` para cada venda.
- `docs/pagina.html` é a página única que todos os clientes usam. Ela lê o
  parâmetro `?id=` da URL, busca o JSON correspondente e monta a timeline.
- `codigos_clientes.csv` guarda a relação `Nº Venda → código`, para que o
  link de um cliente **nunca mude**, mesmo depois de rodar o script várias
  vezes.
- `links_para_enviar.csv` é gerado a cada execução com o link pronto de cada
  cliente, para você copiar e colar no WhatsApp. **Esse arquivo não deve ir
  para o GitHub** (já está no `.gitignore`) porque junta nome do cliente com
  o link de acesso dele.

## 1. Instalar o Python (uma vez só)

Neste computador o Python ainda não está instalado de verdade (só o atalho
da Microsoft Store). Baixe o instalador oficial em
[python.org/downloads](https://www.python.org/downloads/), marque a opção
**"Add python.exe to PATH"** durante a instalação, e depois confirme no
terminal:

```bash
python --version
```

Em seguida instale a única dependência do script:

```bash
pip install openpyxl
```

## 2. Configurar o link do site (uma vez só)

Abra `config.json` e troque a URL de exemplo pela URL real do seu GitHub
Pages (você vai criar essa URL no passo 4). Por enquanto pode deixar como
está — o script funciona normalmente, só o link impresso no final ficará
com o endereço de exemplo até você atualizar.

## 3. Gerar as páginas a partir da planilha

Sempre que atualizar a aba "Pós Venda 2026" no Excel, exporte/salve o
arquivo e rode:

```bash
python gerar_paginas.py "Pós Venda 2026.xlsx"
```

Isso atualiza tudo dentro de `docs/data/` e mostra no terminal a lista de
links prontos (também salva em `links_para_enviar.csv`).

Se o nome da aba for diferente, informe com `--aba`:

```bash
python gerar_paginas.py "Pós Venda 2026.xlsx" --aba "Pós Venda 2027"
```

## 4. Publicar no GitHub Pages (configuração inicial, uma vez só)

1. Crie um repositório novo no GitHub (pode ser público — os dados só ficam
   visíveis para quem tem o link com o código único de cada cliente).
2. Dentro da pasta `PosVenda-IgorLacerda`, rode:

   ```bash
   git init
   git remote add origin https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git
   git add gerar_paginas.py publicar.ps1 config.json .gitignore codigos_clientes.csv README.md docs
   git commit -m "Primeira versão do site de pós-venda"
   git branch -M main
   git push -u origin main
   ```

3. No GitHub, vá em **Settings → Pages**, em "Build and deployment" escolha
   **Deploy from a branch**, selecione a branch `main` e a pasta `/docs`,
   e salve.
4. Depois de ~1 minuto, o GitHub mostra a URL do site (algo como
   `https://seu-usuario.github.io/seu-repositorio/`). Atualize o
   `config.json` com essa URL + `pagina.html` (ex:
   `https://seu-usuario.github.io/seu-repositorio/pagina.html`).
5. O link que você manda pro cliente é sempre:
   `<sua-url>/pagina.html?id=<codigo do cliente>`
   — esse código está em `links_para_enviar.csv` depois de cada execução.

## 5. Atualizações seguintes (o dia a dia)

O script lê direto da planilha mestre do CRM
(`Controle_de_Clientes_Igor_Lacerda_Imoveis_OTIMIZADO_3.xlsx`, aba "Pós Venda
2026") — não precisa manter um arquivo separado. Sempre que adicionar um
cliente novo ou atualizar uma etapa nessa aba, salve o Excel e rode:

```powershell
.\publicar.ps1
```

(sem precisar informar o caminho da planilha — já é o padrão do script).
Esse comando roda o `gerar_paginas.py` e já faz `git add` + `commit` + `push`
automaticamente. Em cerca de 1 minuto o GitHub Pages atualiza sozinho — o
link enviado ao cliente continua o mesmo, só o conteúdo muda.

Se quiser usar outra planilha pontualmente:

```powershell
.\publicar.ps1 "C:\caminho\Outra Planilha.xlsx"
```

Se preferir fazer manualmente:

```bash
python gerar_paginas.py "C:\Users\igorl\Desktop\Controle_de_Clientes_Igor_Lacerda_Imoveis_OTIMIZADO_3.xlsx"
git add docs/data codigos_clientes.csv
git commit -m "Atualiza pós-venda"
git push
```

## Logo

Coloque o arquivo `Igor_Lacerda_MD_1.png` dentro de `docs/assets/`. Se o
arquivo não existir, a página funciona normalmente, só sem a logo no topo.

## Fonte "Ametis"

Ametis não está disponível via CDN público. Se você tiver os arquivos da
fonte (`.woff2`), coloque-os em `docs/assets/fonts/` e descomente o bloco
`@font-face` no início do `<style>` de `pagina.html`. Sem isso, os títulos
usam **Poppins** (via Google Fonts) como substituta visualmente próxima.

## Observações por etapa (opcional)

O script já suporta o campo `observacao` no JSON, mas a planilha atual não
tem colunas de observação por etapa — hoje elas sempre saem em branco. Se no
futuro você quiser adicionar comentários por etapa (ex: "vistoria remarcada
para dia 5"), me avise que ajusto o script para ler colunas extras da
planilha.

## Endereço do imóvel

A planilha atual não tem uma coluna de endereço — a página usa o campo
"Tipo" (Apto, Lote, Sala, Chácara) como identificação do imóvel no
cabeçalho. Se quiser mostrar o endereço completo, adicione uma coluna
"Endereço" na planilha e me avise para eu ajustar o script.

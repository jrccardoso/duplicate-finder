# duplicate-finder

Deteta ficheiros duplicados numa ou mais diretorias (e sub-diretorias), gera um relatório Excel formatado e, opcionalmente, move os duplicados para uma pasta de quarentena.

![Python](https://img.shields.io/badge/python-3.12%2B-blue)
![Platform](https://img.shields.io/badge/plataforma-Windows-lightgrey)

## Funcionalidades

- 🔍 Analisa **múltiplas diretorias** em simultâneo, incluindo sub-pastas
- ⚡ Deteção eficiente em duas fases: agrupa primeiro por **tamanho**, só calcula o **hash MD5** nos candidatos
- 📊 Gera um **relatório Excel** com duas folhas (Resumo + Detalhes) formatadas
- ✍️ **Coluna editável `Manter?`** — escolhe no Excel qual a cópia a preservar em cada grupo
- ✅ Modo **`--apply`** — lê o Excel preenchido e move os não-mantidos, **validando** que cada grupo tem pelo menos um ficheiro marcado com `SIM`
- 📦 Move duplicados para uma **pasta de quarentena** (mantém sempre um original)
- 🛡️ **Modo seguro por omissão** — sem a flag `--move`, nada é movido; apenas gera o relatório

## Requisitos

- Python 3.12 ou superior
- [openpyxl](https://pypi.org/project/openpyxl/) (instalado via `requirements.txt`)

## Instalação

### Windows (rápido)

Executa o script de setup, que cria o ambiente virtual e instala as dependências:

```bash
setup.bat
```

### Manual (multiplataforma)

```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux/macOS:
source venv/bin/activate

pip install -r requirements.txt
```

## Utilização

A ferramenta tem dois modos: **análise** (gera o Excel) e **aplicar** (lê o Excel preenchido e move).

```bash
# Modo análise
python duplicate_finder.py <dir1> [dir2 ...] [--quarantine <pasta>] [--move]

# Modo aplicar (a partir de um Excel revisto)
python duplicate_finder.py --apply <relatorio.xlsx> [--quarantine <pasta>] [--dry-run]
```

| Argumento | Descrição |
|---|---|
| `dir1 [dir2 ...]` | Uma ou mais diretorias a analisar (modo análise) |
| `--quarantine <pasta>` | Pasta de quarentena. Por omissão: `<primeira diretoria>\Quarentena_Duplicados` (análise) ou `<pasta do Excel>\Quarentena_Duplicados` (aplicar) |
| `--move` | Ativa a movimentação dos duplicados já na análise. Sem esta flag, apenas gera o relatório |
| `--apply <excel>` | Lê um relatório Excel já preenchido e move os não-mantidos para quarentena |
| `--dry-run` | Com `--apply`: valida e mostra o que seria movido, **sem mover nada** |

### Fluxo recomendado (revisão humana em 2 passos)

Este fluxo permite-te escolher manualmente qual das cópias fica como "original" antes de mover fosse o que for:

1. **Analisar** — gera o Excel (nada é movido):

   ```bash
   python duplicate_finder.py "C:\Pasta1" "C:\Pasta2"
   ```

2. **Rever no Excel** — abre o `.xlsx`, vai à folha **Detalhes** e, na coluna amarela **`Manter?`**, escreve `SIM` na linha da cópia que queres preservar em cada grupo. A cópia auto-detetada já vem com `SIM` — muda para a que preferires. Podes manter mais do que uma por grupo; tens de manter pelo menos uma.

3. **Aplicar** — lê o Excel e move as restantes para quarentena:

   ```bash
   # Simulação primeiro (recomendado): mostra o que seria movido
   python duplicate_finder.py --apply "C:\Pasta1\relatorio_duplicados_20250101_120000.xlsx" --dry-run

   # Execução real
   python duplicate_finder.py --apply "C:\Pasta1\relatorio_duplicados_20250101_120000.xlsx"
   ```

   Se algum grupo ficar sem nenhum `SIM`, o script **aborta e não move nada**, indicando os grupos por corrigir.

### Outros exemplos

Só gerar o relatório (modo seguro, nada é movido):

```bash
python duplicate_finder.py "C:\Projetos\Cliente"
```

Analisar e mover logo os duplicados (mantém o primeiro de cada grupo, sem revisão manual):

```bash
python duplicate_finder.py "C:\Pasta1" "C:\Pasta2" --quarantine "C:\Quarentena" --move
```

## Como funciona

1. **Scan** — percorre todas as diretorias indicadas e agrupa os ficheiros por tamanho.
2. **Hashing** — apenas nos grupos com mais de um ficheiro do mesmo tamanho calcula o hash MD5, confirmando duplicados reais sem processar ficheiros únicos.
3. **Relatório** — cria um `.xlsx` com o resumo (grupos, ficheiros a remover, espaço desperdiçado) e o detalhe de cada grupo. A folha **Detalhes** inclui a coluna editável `Manter?` e uma coluna `Hash` que identifica os grupos de forma fiável no passo de aplicação.
4. **Quarentena** (opcional) — com `--move`, move os duplicados para a pasta de quarentena, mantendo sempre o primeiro ficheiro de cada grupo como original.
5. **Aplicar** (opcional) — com `--apply <excel>`, relê o relatório preenchido, valida que cada grupo tem pelo menos um `SIM` na coluna `Manter?` e move as cópias não-marcadas para quarentena. Os ficheiros são agrupados pelo `Hash`, pelo que podes reordenar as linhas do Excel à vontade.

## Output

- **Relatório Excel:** `relatorio_duplicados_<timestamp>.xlsx`, criado na primeira diretoria analisada.
- **Quarentena** (com `--move`): pasta `Quarentena_Duplicados` com os ficheiros duplicados movidos. Colisões de nome são resolvidas automaticamente com um sufixo numérico.

## Contribuições

Contribuições são bem-vindas! Abre um _issue_ para reportar bugs ou sugerir melhorias,
ou envia um _pull request_.

## Licença

Distribuído sob a licença MIT. Ver [LICENSE](LICENSE) para mais detalhes.

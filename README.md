# duplicate-finder

Deteta ficheiros duplicados numa ou mais diretorias (e sub-diretorias), gera um relatório Excel formatado e, opcionalmente, move os duplicados para uma pasta de quarentena.

![Python](https://img.shields.io/badge/python-3.12%2B-blue)
![Platform](https://img.shields.io/badge/plataforma-Windows-lightgrey)

## Funcionalidades

- 🔍 Analisa **múltiplas diretorias** em simultâneo, incluindo sub-pastas
- ⚡ Deteção eficiente em duas fases: agrupa primeiro por **tamanho**, só calcula o **hash MD5** nos candidatos
- 📊 Gera um **relatório Excel** com duas folhas (Resumo + Detalhes) formatadas
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

```bash
python duplicate_finder.py <dir1> [dir2 ...] [--quarantine <pasta>] [--move]
```

| Argumento | Descrição |
|---|---|
| `dir1 [dir2 ...]` | Uma ou mais diretorias a analisar (obrigatório) |
| `--quarantine <pasta>` | Pasta de quarentena. Por omissão: `<primeira diretoria>\Quarentena_Duplicados` |
| `--move` | Ativa a movimentação dos duplicados. Sem esta flag, apenas gera o relatório |

### Exemplos

Só gerar o relatório (modo seguro, nada é movido):

```bash
python duplicate_finder.py "C:\Projetos\Cliente"
```

Analisar várias pastas:

```bash
python duplicate_finder.py "C:\Pasta1" "C:\Pasta2"
```

Analisar e mover os duplicados para uma quarentena específica:

```bash
python duplicate_finder.py "C:\Pasta1" "C:\Pasta2" --quarantine "C:\Quarentena" --move
```

## Como funciona

1. **Scan** — percorre todas as diretorias indicadas e agrupa os ficheiros por tamanho.
2. **Hashing** — apenas nos grupos com mais de um ficheiro do mesmo tamanho calcula o hash MD5, confirmando duplicados reais sem processar ficheiros únicos.
3. **Relatório** — cria um `.xlsx` com o resumo (grupos, ficheiros a remover, espaço desperdiçado) e o detalhe de cada grupo.
4. **Quarentena** (opcional) — com `--move`, move os duplicados para a pasta de quarentena, mantendo sempre o primeiro ficheiro de cada grupo como original.

## Output

- **Relatório Excel:** `relatorio_duplicados_<timestamp>.xlsx`, criado na primeira diretoria analisada.
- **Quarentena** (com `--move`): pasta `Quarentena_Duplicados` com os ficheiros duplicados movidos. Colisões de nome são resolvidas automaticamente com um sufixo numérico.

## Licença

Ver [LICENSE](LICENSE). Confirma a licença aplicável antes de publicar o repositório.

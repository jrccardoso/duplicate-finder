"""
duplicate_finder.py
-------------------
Encontra ficheiros duplicados em uma ou mais diretorias (e sub-diretorias),
gera um relatório Excel e move os duplicados para uma pasta de quarentena.

Uso:
    python duplicate_finder.py <dir1> [dir2 ...] [--quarantine <pasta>] [--move]

Exemplos:
    python duplicate_finder.py "C:\\Projetos\\Cliente"
    python duplicate_finder.py "C:\\Pasta1" "C:\\Pasta2"
    python duplicate_finder.py "C:\\Pasta1" "C:\\Pasta2" --quarantine "C:\\Quarentena" --move

Dependências (instalar uma vez):
    pip install openpyxl
"""

import os
import sys
import hashlib
import shutil
import argparse
from collections import defaultdict
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Utilitários ──────────────────────────────────────────────────────────────

def format_size(size_bytes):
    """Converte bytes para formato legível."""
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size_bytes < 1024:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.2f} PB"


def hash_file(path, chunk_size=8192):
    """Calcula o hash MD5 de um ficheiro."""
    h = hashlib.md5()
    try:
        with open(path, "rb") as f:
            while chunk := f.read(chunk_size):
                h.update(chunk)
        return h.hexdigest()
    except (PermissionError, OSError):
        return None


# ── Lógica principal ─────────────────────────────────────────────────────────

def scan_files(root_dirs):
    """Percorre uma ou mais diretorias e agrupa todos os ficheiros por tamanho."""
    size_map = defaultdict(list)
    total = 0

    for root_dir in root_dirs:
        print(f"\n🔍 A analisar: {root_dir}")
        dir_total = 0
        for dirpath, _, filenames in os.walk(root_dir):
            for fname in filenames:
                fpath = os.path.join(dirpath, fname)
                try:
                    size = os.path.getsize(fpath)
                    size_map[size].append(fpath)
                    dir_total += 1
                except (PermissionError, OSError):
                    continue
        print(f"   ✅ {dir_total} ficheiros encontrados")
        total += dir_total

    print(f"\n📁 Total combinado: {total} ficheiro(s) em {len(root_dirs)} diretoria(s)")
    return size_map


def find_duplicates(size_map):
    """
    Agrupa por tamanho primeiro (rápido), depois calcula hash
    apenas nos grupos com mais de 1 ficheiro (eficiente).
    """
    print("⚙️  A calcular hashes dos candidatos a duplicados...")
    hash_map = defaultdict(list)
    candidates = sum(len(v) for v in size_map.values() if len(v) > 1)
    processed = 0

    for paths in size_map.values():
        if len(paths) < 2:
            continue
        for path in paths:
            fhash = hash_file(path)
            if fhash:
                hash_map[fhash].append(path)
            processed += 1
            if processed % 50 == 0:
                print(f"   ... {processed}/{candidates} ficheiros processados")

    duplicates = {h: ps for h, ps in hash_map.items() if len(ps) > 1}
    total_dup_groups = len(duplicates)
    total_dup_files = sum(len(v) - 1 for v in duplicates.values())
    print(f"   ✅ {total_dup_groups} grupos de duplicados | {total_dup_files} ficheiros a remover")
    return duplicates


# ── Relatório Excel ──────────────────────────────────────────────────────────

HEADER_FILL   = PatternFill("solid", start_color="1F4E79")
GROUP_FILL    = PatternFill("solid", start_color="D6E4F0")
ORIGINAL_FILL = PatternFill("solid", start_color="E2EFDA")
DUP_FILL      = PatternFill("solid", start_color="FCE4D6")
HEADER_FONT   = Font(name="Arial", bold=True, color="FFFFFF", size=11)
BODY_FONT     = Font(name="Arial", size=10)
BOLD_FONT     = Font(name="Arial", bold=True, size=10)
CENTER        = Alignment(horizontal="center", vertical="center")
LEFT          = Alignment(horizontal="left", vertical="center", wrap_text=True)
thin_border   = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)


def style_cell(cell, font=None, fill=None, alignment=None, border=None):
    if font:      cell.font      = font
    if fill:      cell.fill      = fill
    if alignment: cell.alignment = alignment
    if border:    cell.border    = border


def create_excel_report(duplicates, output_path, moved_files):
    """Gera o relatório Excel com duas folhas: Resumo e Detalhes."""
    wb = Workbook()

    # ── Folha 1: Resumo ───────────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Resumo"
    ws_sum.sheet_view.showGridLines = False

    # Título
    ws_sum.merge_cells("A1:D1")
    ws_sum["A1"] = "Relatório de Ficheiros Duplicados"
    style_cell(ws_sum["A1"], font=Font(name="Arial", bold=True, size=16, color="1F4E79"), alignment=CENTER)
    ws_sum.row_dimensions[1].height = 35

    ws_sum.merge_cells("A2:D2")
    ws_sum["A2"] = f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    style_cell(ws_sum["A2"], font=Font(name="Arial", size=10, color="595959"), alignment=CENTER)
    ws_sum.row_dimensions[2].height = 20

    ws_sum.append([])  # linha vazia

    # Métricas
    total_groups   = len(duplicates)
    total_dup_files = sum(len(v) - 1 for v in duplicates.values())
    total_wasted   = sum(os.path.getsize(p) for paths in duplicates.values() for p in paths[1:] if os.path.exists(p))

    headers_sum = ["Métrica", "Valor"]
    ws_sum.append(headers_sum)
    for col, h in enumerate(headers_sum, 1):
        cell = ws_sum.cell(row=4, column=col)
        style_cell(cell, font=HEADER_FONT, fill=HEADER_FILL, alignment=CENTER, border=thin_border)

    metrics = [
        ("Grupos de duplicados encontrados", total_groups),
        ("Ficheiros duplicados (a remover)", total_dup_files),
        ("Espaço desperdiçado",              format_size(total_wasted)),
        ("Ficheiros movidos para quarentena", len(moved_files)),
    ]
    for i, (label, value) in enumerate(metrics, 5):
        ws_sum[f"A{i}"] = label
        ws_sum[f"B{i}"] = value
        style_cell(ws_sum[f"A{i}"], font=BODY_FONT, alignment=LEFT,   border=thin_border)
        style_cell(ws_sum[f"B{i}"], font=BOLD_FONT, alignment=CENTER, border=thin_border)

    ws_sum.column_dimensions["A"].width = 42
    ws_sum.column_dimensions["B"].width = 30

    # ── Folha 2: Detalhes ─────────────────────────────────────────────────
    ws = wb.create_sheet("Detalhes")
    ws.sheet_view.showGridLines = False

    col_headers = ["Grupo", "Tipo", "Nome do Ficheiro", "Caminho Completo",
                   "Tamanho", "Data de Modificação", "Estado"]
    ws.append(col_headers)

    for col, h in enumerate(col_headers, 1):
        cell = ws.cell(row=1, column=col)
        cell.value = h
        style_cell(cell, font=HEADER_FONT, fill=HEADER_FILL, alignment=CENTER, border=thin_border)
    ws.row_dimensions[1].height = 22

    row = 2
    for group_num, (fhash, paths) in enumerate(sorted(duplicates.items()), 1):
        for idx, path in enumerate(paths):
            tipo  = "✅ Original" if idx == 0 else "🔁 Duplicado"
            fill  = ORIGINAL_FILL if idx == 0 else DUP_FILL
            fname = os.path.basename(path)

            try:
                size  = os.path.getsize(path) if os.path.exists(path) else 0
                mtime = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d/%m/%Y %H:%M") \
                        if os.path.exists(path) else "—"
            except OSError:
                size, mtime = 0, "—"

            estado = "Movido para quarentena" if path in moved_files else (
                     "Original (mantido)"     if idx == 0 else "Duplicado (não movido)")

            values = [group_num, tipo, fname, path, format_size(size), mtime, estado]
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row, column=col, value=val)
                style_cell(cell, font=BODY_FONT, fill=fill, alignment=LEFT, border=thin_border)

            ws.row_dimensions[row].height = 18
            row += 1

        # Linha separadora entre grupos
        ws.append([])
        row += 1

    # Larguras das colunas
    col_widths = [8, 14, 35, 65, 12, 22, 28]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Freezar cabeçalho
    ws.freeze_panes = "A2"

    wb.save(output_path)
    print(f"\n📊 Relatório Excel guardado em: {output_path}")


# ── Quarentena ───────────────────────────────────────────────────────────────

def move_to_quarantine(duplicates, quarantine_dir):
    """Move os duplicados (mantém sempre o primeiro) para a pasta de quarentena."""
    os.makedirs(quarantine_dir, exist_ok=True)
    moved = []

    print(f"\n📦 A mover duplicados para: {quarantine_dir}")
    for paths in duplicates.values():
        for path in paths[1:]:  # O índice 0 é o "original" — fica no lugar
            try:
                dest_name = os.path.basename(path)
                dest      = os.path.join(quarantine_dir, dest_name)
                # Evita colisões de nome na quarentena
                if os.path.exists(dest):
                    base, ext = os.path.splitext(dest_name)
                    dest = os.path.join(quarantine_dir, f"{base}_{len(moved)}{ext}")
                shutil.move(path, dest)
                moved.append(path)
                print(f"   ↪ {os.path.basename(path)}")
            except (PermissionError, OSError) as e:
                print(f"   ⚠️  Erro ao mover {path}: {e}")

    print(f"   ✅ {len(moved)} ficheiro(s) movido(s)")
    return moved


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Detetor de ficheiros duplicados com relatório Excel e quarentena."
    )
    parser.add_argument("roots", nargs="+",
        help="Uma ou mais diretorias a analisar. Ex: C:\\Pasta1 C:\\Pasta2")
    parser.add_argument("--quarantine", default=None,
        help="Pasta de quarentena. Por omissão: <primeira diretoria>\\Quarentena_Duplicados")
    parser.add_argument("--move", action="store_true",
        help="Ativa a movimentação dos duplicados para quarentena. "
             "Sem esta flag só gera o relatório (modo seguro).")

    args = parser.parse_args()
    roots = [os.path.abspath(r) for r in args.roots]

    for r in roots:
        if not os.path.isdir(r):
            print(f"❌ Diretoria não encontrada: {r}")
            sys.exit(1)

    quarantine_dir = args.quarantine or os.path.join(roots[0], "Quarentena_Duplicados")
    timestamp      = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path    = os.path.join(roots[0], f"relatorio_duplicados_{timestamp}.xlsx")

    # 1. Scan
    size_map   = scan_files(roots)

    # 2. Duplicados
    duplicates = find_duplicates(size_map)

    if not duplicates:
        print("\n✅ Nenhum ficheiro duplicado encontrado!")
        sys.exit(0)

    # 3. Mover (opcional)
    moved_files = []
    if args.move:
        moved_files = move_to_quarantine(duplicates, quarantine_dir)
    else:
        print("\nℹ️  Modo de leitura — nenhum ficheiro foi movido.")
        print("   Para mover os duplicados, adiciona --move ao comando.")

    # 4. Relatório Excel
    create_excel_report(duplicates, report_path, moved_files)

    print("\n✅ Concluído!")
    if args.move:
        print(f"   Quarentena : {quarantine_dir}")
    print(f"   Relatório  : {report_path}\n")


if __name__ == "__main__":
    main()

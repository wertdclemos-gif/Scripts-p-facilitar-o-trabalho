from __future__ import annotations

import re
import traceback
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Set

from pypdf import PdfReader, PdfWriter


INPUT_DIR = Path(r"C:\Users\pdv\Desktop\Fechamento_Mensal\VR")
OUTPUT_DIR = Path(r"C:\Users\pdv\Desktop\Fechamento_Mensal\VR\VR SAIDA")


def ensure_dirs() -> None:
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"Pasta de entrada não encontrada: {INPUT_DIR}")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def list_pdf_files(folder: Path) -> List[Path]:
    return sorted(
        [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    )


def extract_common_name(filename: str) -> Optional[str]:
    """
    Extrai o nome em maiúsculas que aparece no final, após hífen '-' ou underline '_'.

    Exemplos:
    - NFSe_59871202_99105374 - ABIN.pdf -> ABIN
    - name - ABIN.pdf -> ABIN
    - relatorio_pedido_departamento_3338491_ABIN.pdf -> ABIN
    """
    stem = Path(filename).stem.strip()

    match_hyphen = re.search(r"-\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,})\s*$", stem)
    if match_hyphen:
        return match_hyphen.group(1).strip()

    match_underscore = re.search(r"_([A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,})\s*$", stem)
    if match_underscore:
        return match_underscore.group(1).strip()

    return None


def group_pdfs(pdf_files: List[Path]) -> Dict[str, List[Path]]:
    groups: Dict[str, List[Path]] = defaultdict(list)

    for pdf in pdf_files:
        common_name = extract_common_name(pdf.name)
        if common_name:
            groups[common_name].append(pdf)

    return groups


def classify_file(pdf_path: Path) -> str:
    """
    Classifica o arquivo em um dos 3 tipos esperados.
    """
    stem = pdf_path.stem.strip()

    if re.match(r"(?i)^NFSe_\d+_\d+\s*-\s*[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,}$", stem):
        return "nfse"

    if re.match(r"(?i)^name\s*-\s*[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,}$", stem):
        return "name"

    if re.match(r"(?i)^relatorio_pedido_departamento_\d+_[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,}$", stem):
        return "relatorio"

    return "outro"


def get_file_order(pdf_path: Path) -> tuple[int, str]:
    """
    Ordem exigida:
    1. NFSe_<inteiros>_<inteiros> - <NOME>
    2. name - <NOME>
    3. relatorio_pedido_departamento_<inteiros>_<NOME>
    """
    kind = classify_file(pdf_path)
    stem = pdf_path.stem.strip()

    if kind == "nfse":
        return (1, stem)
    if kind == "name":
        return (2, stem)
    if kind == "relatorio":
        return (3, stem)

    return (99, stem)


def has_all_required_types(files: List[Path]) -> bool:
    found_types: Set[str] = {classify_file(f) for f in files}
    required = {"nfse", "name", "relatorio"}
    return required.issubset(found_types)


def merge_pdfs(pdf_paths: List[Path], output_path: Path) -> None:
    writer = PdfWriter()

    for pdf_path in pdf_paths:
        reader = PdfReader(str(pdf_path))
        for page in reader.pages:
            writer.add_page(page)

    with output_path.open("wb") as f:
        writer.write(f)


def main() -> None:
    ensure_dirs()

    pdf_files = list_pdf_files(INPUT_DIR)

    if not pdf_files:
        print(f"Nenhum PDF encontrado em: {INPUT_DIR}")
        return

    groups = group_pdfs(pdf_files)

    if not groups:
        print("Nenhum grupo válido foi encontrado.")
        return

    print("Grupos encontrados:")
    for common_name, files in groups.items():
        print(f" - {common_name}: {len(files)} arquivo(s)")

    for common_name, files in groups.items():
        print(f"\nAnalisando grupo: {common_name}")
        for f in files:
            print(f" - {f.name} [{classify_file(f)}]")

        if not has_all_required_types(files):
            print(f" => Grupo ignorado: faltam um ou mais tipos obrigatórios para {common_name}")
            continue

        ordered_files = sorted(files, key=get_file_order)

        output_name = f"VR {common_name}.pdf"
        output_path = OUTPUT_DIR / output_name

        print("Ordem de união:")
        for f in ordered_files:
            print(f" - {f.name}")

        merge_pdfs(ordered_files, output_path)
        print(f" => Gerado: {output_path.name}")

    print("\nConcluído com sucesso.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\nERRO AO EXECUTAR O SCRIPT:")
        print(str(e))
        print("\nDetalhes técnicos:")
        traceback.print_exc()

    input("\nPressione ENTER para fechar...")
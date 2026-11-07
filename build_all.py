# -*- coding: utf-8 -*-
"""Полная пересборка проекта одной командой:  python build_all.py

Порядок: основная книга → пересчёт LibreOffice → книга дополнения → пересчёт → презентация (PPTX → PDF)
→ пояснительная записка (два прохода, чтобы проставить номера страниц в оглавлении) → дополнение к записке.
Нужен LibreOffice (soffice). Путь можно задать переменной окружения SOFFICE.
"""
import glob
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.abspath(__file__))
SOFFICE = os.environ.get("SOFFICE") or shutil.which("soffice") or r"C:\Program Files\LibreOffice\program\soffice.exe"
PROFILE = os.path.join(tempfile.gettempdir(), "fmjc_lo_profile")
PY = sys.executable
ENV = dict(os.environ, PYTHONIOENCODING="utf-8")


def run(args, cwd=ROOT):
    print("→", " ".join(os.path.basename(a) if i == 0 else a for i, a in enumerate(args)))
    subprocess.run(args, cwd=cwd, check=True, env=ENV)


def soffice(src, fmt, outdir):
    os.makedirs(outdir, exist_ok=True)
    url = "file:///" + PROFILE.replace("\\", "/")
    subprocess.run([SOFFICE, "--headless", "--norestore", f"-env:UserInstallation={url}", "--convert-to", fmt,
                    "--outdir", outdir, src], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                   timeout=300)
    return os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + "." + fmt.split(":")[0])


def recalc(xlsx):
    """Пересчитать формулы: LibreOffice открывает книгу и сохраняет её с вычисленными значениями."""
    tmp = tempfile.mkdtemp()
    out = soffice(os.path.abspath(xlsx), "xlsx", tmp)
    shutil.copyfile(out, xlsx)


def main():
    tmp = tempfile.mkdtemp()
    run([PY, "build_workbook.py", "Расчёт_водоотлив.xlsx"])
    recalc(os.path.join(ROOT, "Расчёт_водоотлив.xlsx"))

    add = os.path.join(ROOT, "Дополнение")
    run([PY, "build_addendum_xlsx.py"], cwd=add)
    recalc(os.path.join(add, "Дополнение_расчёты.xlsx"))

    run([PY, "build_presentation.py"])
    for pptx in glob.glob(os.path.join(ROOT, "*_Карьерная_вода_без_остановки.pptx")):
        pdf = soffice(pptx, "pdf", tmp)
        shutil.copyfile(pdf, os.path.splitext(pptx)[0] + ".pdf")

    report = "Пояснительная_записка_Карьерная_вода_без_остановки.docx"
    run([PY, "build_report.py"])
    toc = os.path.join(ROOT, "рисунки_записки", "toc.json")
    pages = os.path.join(tmp, "pages.json")
    run([PY, "find_pages.py", soffice(os.path.join(ROOT, report), "pdf", tmp), toc, pages])
    run([PY, "build_report.py", pages])

    # принципиальная гидравлическая схема (лист A3 + перечень, слайд 5а); нужна до документа дополнения
    run([PY, "build_scheme.py"], cwd=add)
    slide5a = os.path.join(add, "Слайд_5а_гидравлическая_схема.pptx")
    shutil.copyfile(soffice(slide5a, "pdf", tmp), os.path.splitext(slide5a)[0] + ".pdf")

    run([PY, "build_addendum_docx.py"], cwd=add)
    print("Готово.")


if __name__ == "__main__":
    main()

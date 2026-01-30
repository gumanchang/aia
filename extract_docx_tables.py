from pathlib import Path

import pandas as pd
from docx import Document


def extract_tables(docx_path: str | Path, output_path: str | Path | None = None) -> Path:
    """
    从 DOCX 中抽取所有表格，合并导出为一个 Excel 文件。

    :param docx_path: Word 文档路径
    :param output_path: 输出 Excel 路径（默认与 DOCX 同目录下，文件名加 _tables.xlsx）
    :return: 输出文件路径
    """
    docx_path = Path(docx_path).expanduser().resolve()
    if output_path is None:
        output_path = docx_path.with_suffix("").with_name(docx_path.stem + "_tables.xlsx")
    else:
        output_path = Path(output_path).expanduser().resolve()

    document = Document(str(docx_path))

    tables_dfs: list[pd.DataFrame] = []

    for ti, table in enumerate(document.tables, start=1):
        rows = []
        for row in table.rows:
            rows.append([cell.text.strip() for cell in row.cells])

        if not rows:
            continue

        # 简单处理：第一行当表头，如果长度不一致，自动补空
        max_len = max(len(r) for r in rows)
        norm_rows = [r + [""] * (max_len - len(r)) for r in rows]

        header = norm_rows[0]
        data = norm_rows[1:] or [[]]

        df = pd.DataFrame(data, columns=header)
        # 给每张表加上来源标记
        df.insert(0, "表序号", ti)
        tables_dfs.append(df)

    if not tables_dfs:
        raise SystemExit("文档中未检测到任何表格。")

    # 写入 Excel，每个表一个 sheet，同时再合并一个总表
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for idx, df in enumerate(tables_dfs, start=1):
            sheet_name = f"Table_{idx}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        merged = pd.concat(tables_dfs, ignore_index=True)
        merged.to_excel(writer, sheet_name="ALL_TABLES", index=False)

    return output_path


if __name__ == "__main__":
    docx_file = Path("AI工具在软件开发生命周期中的应用研究.docx")
    if not docx_file.exists():
        raise SystemExit(f"未找到 DOCX 文件: {docx_file}")

    out = extract_tables(docx_file)
    print(f"表格已导出到: {out}")


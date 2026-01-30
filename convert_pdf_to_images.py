from pathlib import Path

from pdf2image import convert_from_path


def pdf_to_images(pdf_path: str | Path, output_dir: str | Path | None = None, dpi: int = 200) -> list[Path]:
    """
    将 PDF 每一页导出为 PNG 图片。

    :param pdf_path: PDF 文件路径
    :param output_dir: 输出目录（默认与 PDF 同目录下的 <pdf_name>_pages 文件夹）
    :param dpi: 导出分辨率
    :return: 生成的图片路径列表
    """
    pdf_path = Path(pdf_path).expanduser().resolve()
    if output_dir is None:
        output_dir = pdf_path.with_suffix("").name + "_pages"
        output_dir = pdf_path.parent / output_dir
    else:
        output_dir = Path(output_dir).expanduser().resolve()

    output_dir.mkdir(parents=True, exist_ok=True)

    # 需要本机安装 poppler，并在 PATH 中可用。
    images = convert_from_path(str(pdf_path), dpi=dpi)

    result_paths: list[Path] = []
    for i, image in enumerate(images, start=1):
        out_path = output_dir / f"page_{i:02d}.png"
        image.save(out_path, "PNG")
        result_paths.append(out_path)

    return result_paths


if __name__ == "__main__":
    pdf_file = Path("AI 工具在 SDLC 调研报告.pdf")
    if not pdf_file.exists():
        raise SystemExit(f"未找到 PDF 文件: {pdf_file}")

    generated = pdf_to_images(pdf_file)
    print("生成图片文件：")
    for p in generated:
        print(p)


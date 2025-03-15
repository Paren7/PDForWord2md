import os
from pathlib import Path
import pypandoc
from marker.config.parser import ConfigParser
from marker.converters.pdf import PdfConverter
from marker.models import create_model_dict
from marker.output import text_from_rendered

config = {
    "force_ocr": True,  # 强制 OCR
    "output_format": "markdown",  # 输出格式为 Markdown
    "gpu": True
}

# 创建配置解析器
config_parser = ConfigParser(config)

# 创建 PdfConverter 实例
converter = PdfConverter(
    artifact_dict=create_model_dict(),  # 加载模型
    config=config_parser.generate_config_dict(),  # 加载配置
)


def convert_docx_to_md(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 遍历输入文件夹中的所有文件
    for filename in os.listdir(input_folder):
        if filename.endswith(".docx") or filename.endswith('.doc'):
            # 构建输入和输出文件的完整路径
            input_path = os.path.join(input_folder, filename)
            output_filename = os.path.splitext(filename)[0] + ".md"
            output_path = os.path.join(output_folder, output_filename)

            # 使用 pandoc 进行转换
            print(f"Converting {input_path} to {output_path}")
            pypandoc.convert_file(input_path, 'md', outputfile=output_path)


def convert_pdf_to_markdown(pdf_path, output_folder, force_ocr=True, use_gpu=True):
    rendered = converter(str(pdf_path))
    text, _, images = text_from_rendered(rendered)
    output_path = Path(output_folder) / f"{Path(pdf_path).stem}.md"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)

    print(f"Converted {pdf_path} to {output_path}")


def convert_folder_pdfs_to_markdown(input_folder, output_folder, force_ocr=True, use_gpu=True):
    # 确保输出文件夹存在
    Path(output_folder).mkdir(parents=True, exist_ok=True)

    # 遍历输入文件夹中的所有 PDF 文件
    for pdf_file in Path(input_folder).glob("*.pdf"):
        convert_pdf_to_markdown(pdf_file, output_folder, force_ocr, use_gpu)


if __name__ == "__main__":
    input_folder = "F:\project\PyT\pythonProject\yjs"

    output_folder = "F:\project\PyT\pythonProject\yjs/output"
    convert_docx_to_md(input_folder, output_folder)
    convert_folder_pdfs_to_markdown(input_folder, output_folder)

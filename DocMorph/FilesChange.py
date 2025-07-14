import os
from typing import Optional
from pdf2docx import Converter
import pypandoc
import pandas as pd
from docx2pdf import convert as docx2pdf_convert
from PyPDF2 import PdfReader, PdfWriter
import glob
from concurrent.futures import ThreadPoolExecutor, as_completed


class DocumentConverter:
    # 确保路径创建
    os.makedirs('./input', exist_ok=True)
    os.makedirs('./output', exist_ok=True)

    def __init__(self):
        """初始化支持的文档转换类型"""
        # 定义支持的转换格式映射
        self.supported_conversions = {
            'pdf': ['docx', 'txt'],  # PDF可转换为Word或纯文本
            'docx': ['pdf', 'md', 'html'],  # Word可转换为PDF/Markdown/HTML
            'xlsx': ['csv'],  # Excel可转换为CSV
            'md': ['docx', 'html'],  # Markdown可转换为Word或HTML
            'html': ['docx', 'pdf', 'md']  # HTML可转换为Word/PDF/Markdown
        }


    def get_supported_conversions(self, from_format: str) -> list:
        """获取某格式支持转换的目标格式列表
        :param from_format: 源文件格式
        :return: 支持的目标格式列表
        """
        return self.supported_conversions.get(from_format.lower(), [])

    def convert(self, input_path: str, output_path: str,
                from_format: Optional[str] = None,
                to_format: Optional[str] = None) -> bool:
        """
        执行文档转换的主方法
        :param input_path: 输入文件路径
        :param output_path: 输出文件路径
        :param from_format: 可选，手动指定输入文件格式
        :param to_format: 可选，手动指定输出文件格式
        :return: 转换是否成功
        """
        # 自动检测格式（如果未手动指定）
        if from_format is None:
            from_format = os.path.splitext(input_path)[1][1:].lower()  # 从扩展名获取格式
        if to_format is None:
            to_format = os.path.splitext(output_path)[1][1:].lower()

        # 检查是否支持该转换
        if to_format not in self.get_supported_conversions(from_format):
            raise ValueError(f"不支持从 {from_format} 转换为 {to_format}")

        # 根据格式选择对应的转换方法
        try:
            if from_format == 'pdf' and to_format == 'docx':
                self._pdf_to_docx(input_path, output_path)
            elif from_format == 'docx' and to_format == 'pdf':
                self._docx_to_pdf(input_path, output_path)
            elif from_format == 'xlsx' and to_format == 'csv':
                self._excel_to_csv(input_path, output_path)
            elif from_format == 'docx' and to_format in ['md', 'html']:
                self._docx_to_other(input_path, output_path, to_format)
            elif from_format == 'md' and to_format == 'docx':
                self._markdown_to_docx(input_path, output_path)
            elif from_format == 'html' and to_format == 'pdf':
                self._html_to_pdf(input_path, output_path)
            else:
                # 使用pandoc作为后备转换器
                self._convert_with_pandoc(input_path, output_path, from_format, to_format)
            return True
        except Exception as e:
            print(f"转换失败: {e}")
            return False

    def _pdf_to_docx(self, input_path: str, output_path: str):
        """内部方法：PDF转Word文档"""
        cv = Converter(input_path)
        cv.convert(output_path)
        cv.close()

    def _docx_to_pdf(self, input_path: str, output_path: str):
        # """内部方法：Word文档转PDF (有概率失败)"""
        # docx2pdf_convert(input_path, output_path)

        # pypandoc来处理.docx ->.pdf
        pypandoc.convert_file(input_path, 'pdf', outputfile=output_path)

    def _excel_to_csv(self, input_path: str, output_path: str):
        """内部方法：Excel转CSV"""
        data = pd.read_excel(input_path)
        data.to_csv(output_path, index=False)

    def _docx_to_other(self, input_path: str, output_path: str, to_format: str):
        """内部方法：Word转Markdown或HTML"""
        pypandoc.convert_file(input_path, to_format, outputfile=output_path)

    def _markdown_to_docx(self, input_path: str, output_path: str):
        """内部方法：Markdown转Word"""
        pypandoc.convert_file(input_path, 'docx', outputfile=output_path)

    def _html_to_pdf(self, input_path: str, output_path: str):
        """内部方法：HTML转PDF"""
        pypandoc.convert_file(input_path, 'pdf', outputfile=output_path)

    def _convert_with_pandoc(self, input_path: str, output_path: str,
                             from_format: str, to_format: str):
        """内部方法：使用pandoc进行通用转换"""
        pypandoc.convert_file(input_path, to_format, outputfile=output_path)

    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """从PDF提取纯文本内容
        :param pdf_path: PDF文件路径
        :return: 提取的文本字符串
        """
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        return text

    def merge_pdfs(self, pdf_paths: list, output_path: str):
        """合并多个PDF文件
        :param pdf_paths: PDF文件路径列表
        :param output_path: 合并后的输出路径
        """
        merger = PdfWriter()
        for pdf in pdf_paths:
            merger.append(pdf)
        merger.write(output_path)
        merger.close()

# 批量装换
    def batch_convert(self, input_dir: str, output_dir: str,
                      from_format: str, to_format: str,
                      keep_structure: bool = False,
                      max_workers: int =5 ) -> int:
        """
        批量转换目录下的文件
        :param input_dir: 输入目录
        :param output_dir: 输出目录
        :param from_format: 原始格式（如 'docx'）
        :param to_format: 目标格式（如 'pdf'）
        :param keep_structure: 是否保留原始目录结构
        :return: 成功转换的文件数

        max_workers--使用线程池并发批量转换文件,最大线程
        """
        from_format = from_format.lower()
        to_format = to_format.lower()

        if to_format not in self.get_supported_conversions(from_format):
            raise ValueError(f"不支持从 {from_format} 转换为 {to_format}")

        os.makedirs(output_dir, exist_ok=True)

        files = glob.glob(os.path.join(input_dir, f"**/*.{from_format}"), recursive=True)
        if not files:
            print("❗未找到待转换文件")
            return 0

        def convert_file(file):
            try:
                rel_path = os.path.relpath(file, input_dir)
                new_name = os.path.splitext(rel_path)[0] + f".{to_format}"

                if keep_structure:
                    out_path = os.path.join(output_dir, new_name)
                    os.makedirs(os.path.dirname(out_path), exist_ok=True)
                else:
                    out_path = os.path.join(output_dir, os.path.basename(new_name))

                self.convert(file, out_path, from_format, to_format)
                return True
            except Exception as e:
                print(f"❌ 跳过文件 {file}，原因: {e}")
                return False

        success_count = 0
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(convert_file, file) for file in files]
            for f in as_completed(futures):
                if f.result():
                    success_count += 1

        return success_count



# main函数
def main():
    parser = argparse.ArgumentParser(description="文档格式转换工具")
    parser.add_argument('input', help="输入文件或目录路径" ,default='./input')
    parser.add_argument('output', help="输出文件或目录路径" ,default='./output')
    parser.add_argument('--from', dest='from_format', help="源格式（必须在批量转换中指定）")
    parser.add_argument('--to', dest='to_format', help="目标格式")
    parser.add_argument('--batch', action='store_true', help="启用批量模式")

    args = parser.parse_args()
    converter = DocumentConverter()

    if args.batch:
        if not args.from_format or not args.to_format:
            print("批量模式下必须指定 --from 和 --to 参数")
            return
        count = converter.batch_convert(args.input, args.output,
                                        args.from_format, args.to_format)
        print(f"✅ 批量转换完成，共成功转换 {count} 个文件")
    else:
        success = converter.convert(args.input, args.output,
                                    args.from_format, args.to_format)
        if success:
            print(f"✅ 转换成功: {args.input} -> {args.output}")
        else:
            print("❌ 转换失败")

# 启动
if __name__ == '__main__':
    import argparse  # 移到顶部更合适，但为了不改变原有结构保留在此

    main()
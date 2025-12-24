import os
import re
import argparse
from pathlib import Path
import PyPDF2
import pdfplumber


def extract_year_from_text(text):
    """从文本中提取年份"""
    # 常见的年份模式：4位数字，通常在1900-2100之间
    year_patterns = [
        r'\b(19[0-9]{2}|20[0-2][0-9])\b',  # 1900-2029
        r'\((\d{4})\)',  # (2023) 格式
        r'\b(\d{4})\s*[,-]?\s*(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b',
        r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*[,-]?\s*(\d{4})\b',
    ]

    for pattern in year_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        if matches:
            # 返回找到的第一个有效年份
            for match in matches:
                if isinstance(match, tuple):
                    match = match[0]  # 如果是分组匹配，取第一个分组
                year = int(match)
                if 1900 <= year <= 2030:  # 合理的年份范围
                    return str(year)
    return None


def extract_year_from_pdf(pdf_path):
    """从PDF中提取年份"""
    try:
        # 方法1: 从元数据中提取
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            metadata = pdf_reader.metadata

            if metadata:
                # 检查创建日期和修改日期
                date_fields = ['/CreationDate', '/ModDate']
                for field in date_fields:
                    if field in metadata:
                        date_str = metadata[field]
                        # PDF日期格式: D:YYYYMMDDHHMMSS
                        year_match = re.search(r'D:(\d{4})', date_str)
                        if year_match:
                            year = year_match.group(1)
                            if 1900 <= int(year) <= 2030:
                                return year

        # 方法2: 从文本内容中提取
        with pdfplumber.open(pdf_path) as pdf:
            # 检查前3页
            for page_num in range(min(3, len(pdf.pages))):
                page = pdf.pages[page_num]
                text = page.extract_text()

                if text:
                    # 在文本中搜索年份
                    year = extract_year_from_text(text)
                    if year:
                        return year

                    # 特别检查页眉页脚区域（年份常出现在这里）
                    try:
                        # 检查页面顶部区域（页眉）
                        top_region = page.within_bbox((0, 0, page.width, page.height * 0.2))
                        top_text = top_region.extract_text()
                        if top_text:
                            year = extract_year_from_text(top_text)
                            if year:
                                return year

                        # 检查页面底部区域（页脚）
                        bottom_region = page.within_bbox((0, page.height * 0.8, page.width, page.height))
                        bottom_text = bottom_region.extract_text()
                        if bottom_text:
                            year = extract_year_from_text(bottom_text)
                            if year:
                                return year
                    except:
                        pass  # 如果区域提取失败，继续其他方法

    except Exception as e:
        print(f"提取年份失败 {pdf_path}: {e}")

    return None


def extract_title_with_pypdf2(pdf_path):
    """使用PyPDF2提取标题（从元数据）"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            metadata = pdf_reader.metadata
            if metadata and '/Title' in metadata:
                title = metadata['/Title']
                if title and title.strip():
                    return title.strip()
    except Exception as e:
        print(f"PyPDF2提取标题失败 {pdf_path}: {e}")
    return None


def extract_title_with_pdfplumber(pdf_path):
    """使用pdfplumber从文本内容中提取标题"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 检查第一页
            first_page = pdf.pages[0]
            text = first_page.extract_text()

            if not text:
                return None

            # 清理文本
            lines = [line.strip() for line in text.split('\n') if line.strip()]

            # 寻找可能的标题行（通常在第一页的前几行）
            potential_titles = []
            for i, line in enumerate(lines[:10]):  # 检查前10行
                # 标题通常特征：较短、不含页码、不含特定关键词
                if (len(line) > 10 and len(line) < 200 and
                        not re.search(r'abstract|introduction|references|page|\d{1,2}\s*$', line.lower()) and
                        not re.search(r'^[0-9\s\.\-]*$', line)):
                    potential_titles.append((i, line))

            # 优先选择靠前的、长度适中的行作为标题
            if potential_titles:
                # 按位置排序，选择最靠前的
                potential_titles.sort(key=lambda x: x[0])
                return potential_titles[0][1]

    except Exception as e:
        print(f"pdfplumber提取标题失败 {pdf_path}: {e}")
    return None


def extract_title_advanced(pdf_path):
    """更智能的标题提取方法"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text()

            if not text:
                return None

            lines = [line.strip() for line in text.split('\n') if line.strip()]

            # 排除明显的非标题行
            excluded_keywords = [
                'abstract', 'introduction', 'keywords', 'reference',
                'journal', 'vol', 'volume', 'pp', 'page', 'doi',
                'proceedings', 'conference', 'university', 'department'
            ]

            for i, line in enumerate(lines[:15]):  # 检查前15行
                line_lower = line.lower()

                # 排除条件
                if (len(line) < 10 or len(line) > 250 or
                        any(keyword in line_lower for keyword in excluded_keywords) or
                        re.search(r'^\d{1,4}\s*$', line) or  # 纯数字
                        re.search(r'^[ivxlc]+$', line, re.IGNORECASE) or  # 罗马数字
                        re.search(r'^[a-z]\s*$', line) or  # 单个字母
                        re.search(r'\.{3,}', line) or  # 多个点
                        line.count('.') > 5):  # 太多点号
                    continue

                # 标题特征：通常包含大写字母、可能有标点但不会太多
                if (re.search(r'[A-Z]', line) and  # 包含大写字母
                        line.count('.') <= 3 and  # 点号不多
                        not line.endswith('.') and  # 不以点号结尾（可能是句子）
                        not line.startswith('Received') and  # 排除特定开头
                        not line.startswith('Copyright')):

                    # 进一步清理：移除可能的作者名（如果有逗号分隔且看起来像名字）
                    if ',' in line and len(line.split(',')) <= 3:
                        # 可能是"姓, 名"格式，跳过
                        continue

                    return line

    except Exception as e:
        print(f"高级标题提取失败 {pdf_path}: {e}")
    return None


def sanitize_filename(title):
    """清理标题，使其适合作为文件名"""
    if not title:
        return None

    # 移除或替换非法字符
    illegal_chars = r'[<>:"/\\|?*]'
    title = re.sub(illegal_chars, '', title)

    # 移除多余空格
    title = re.sub(r'\s+', ' ', title).strip()

    # 限制长度
    if len(title) > 100:  # 稍微缩短，因为要加上年份前缀
        title = title[:100] + "..."

    return title


def rename_pdf_files(folder_path, dry_run=True):
    """重命名文件夹中的PDF文件"""
    folder = Path(folder_path)
    pdf_files = list(folder.glob("*.pdf"))

    if not pdf_files:
        print("未找到PDF文件")
        return

    print(f"找到 {len(pdf_files)} 个PDF文件")

    renamed_count = 0
    failed_files = []

    for pdf_file in pdf_files:
        print(f"\n处理文件: {pdf_file.name}")

        # 尝试多种方法提取标题
        title = None
        methods = [
            ("元数据提取", extract_title_with_pypdf2),
            ("内容分析", extract_title_with_pdfplumber),
            ("智能识别", extract_title_advanced)
        ]

        for method_name, method_func in methods:
            title = method_func(pdf_file)
            if title:
                print(f"  {method_name}成功: {title[:80]}...")
                break
            else:
                print(f"  {method_name}失败")

        if not title:
            print(f"  无法提取标题，跳过此文件")
            failed_files.append(pdf_file.name)
            continue

        # 清理标题
        clean_title = sanitize_filename(title)
        if not clean_title:
            print(f"  标题清理失败，跳过此文件")
            failed_files.append(pdf_file.name)
            continue

        # 提取年份
        year = extract_year_from_pdf(pdf_file)
        if year:
            print(f"  识别到年份: {year}")
        else:
            print(f"  未识别到年份，使用'未知年份'")
            year = "未知年份"

        # 生成新文件名：年份_标题
        new_filename = f"{year}_{clean_title}.pdf"
        new_filepath = pdf_file.parent / new_filename

        # 检查是否已存在同名文件
        counter = 1
        original_new_filepath = new_filepath
        while new_filepath.exists() and new_filepath != pdf_file:
            new_filename = f"{year}_{clean_title}_{counter}.pdf"
            new_filepath = pdf_file.parent / new_filename
            counter += 1

        if dry_run:
            print(f"  预览重命名: {pdf_file.name} -> {new_filename}")
        else:
            try:
                pdf_file.rename(new_filepath)
                print(f"  成功重命名: {new_filename}")
                renamed_count += 1
            except Exception as e:
                print(f"  重命名失败: {e}")
                failed_files.append(pdf_file.name)

    print(f"\n{'=' * 50}")
    print(f"处理完成!")
    if dry_run:
        print(f"预览模式 - 实际会重命名 {renamed_count} 个文件")
    else:
        print(f"成功重命名 {renamed_count} 个文件")

    if failed_files:
        print(f"失败文件 ({len(failed_files)} 个):")
        for failed in failed_files:
            print(f"  - {failed}")


def main():
    parser = argparse.ArgumentParser(description="智能重命名PDF论文文件")
    parser.add_argument("folder", help="包含PDF文件的文件夹路径")
    parser.add_argument("--dry-run", action="store_true",
                        help="预览模式，不实际重命名文件")

    args = parser.parse_args()

    if not os.path.exists(args.folder):
        print(f"错误: 文件夹 '{args.folder}' 不存在")
        return

    rename_pdf_files(args.folder, dry_run=args.dry_run)


if __name__ == "__main__":
    # 直接在这里指定路径
    folder_path = r"F:\BaiduSyncdisk-mu\Projects CUG\毕业论文\250616_调研_水的要素加入"  # Windows

    # 先预览
    rename_pdf_files(folder_path, dry_run=True)

    # 确认后实际执行
    confirm = input("确认要执行重命名吗？(y/n): ")
    if confirm.lower() == 'y':
        rename_pdf_files(folder_path, dry_run=False)
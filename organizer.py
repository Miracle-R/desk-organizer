import shutil
import csv
import yaml
from datetime import datetime
from pathlib import Path

import docx
import pdfplumber
from pptx import Presentation
import openpyxl
import magic

# ========== 1. 读取分类配置 ==========
def load_categories(config_path='categories.yaml'):
    with open(config_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f)
    return data['categories']

# ========== 2. 从文档中提取文本（全格式支持） ==========
def extract_text(file_path):
    suffix = file_path.suffix.lower()
    try:
        mime = magic.Magic(mime=True).from_file(str(file_path))
    except Exception:
        mime = ''
    try:
        if suffix in ('.docx', '.wps') or 'wordprocessingml' in mime:
            doc = docx.Document(file_path)
            return '\n'.join([p.text for p in doc.paragraphs])
        if suffix in ('.xlsx', '.et') or 'spreadsheetml' in mime:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            texts = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    row_text = ' '.join([str(c) for c in row if c is not None])
                    if row_text.strip():
                        texts.append(row_text)
            return '\n'.join(texts)
        if suffix in ('.pptx', '.dps') or 'presentationml' in mime:
            prs = Presentation(file_path)
            texts = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, 'text'):
                        texts.append(shape.text)
            return '\n'.join(texts)
        if suffix == '.pdf' or 'pdf' in mime:
            text = ''
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages[:3]:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + '\n'
            return text
        if suffix in ('.txt', '.md'):
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        if suffix == '.csv':
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                reader = csv.reader(f)
                rows = [' '.join(row) for row in reader if any(row)]
                return '\n'.join(rows)
        return ''
    except Exception as e:
        print(f'  [!] 读取失败 {file_path.name}: {e}')
        return ''

# ========== 3. 基于标题优先的分类 ==========
def classify_document(text, categories):
    title_region = text[:200].lower()
    full_text = text.lower()
    scores = {}
    for cat in categories:
        score = 0
        tit_kws = cat.get('title_keywords', [])
        full_kws = cat.get('fulltext_keywords', [])
        for kw in tit_kws:
            score += title_region.count(kw) * 3
        for kw in full_kws:
            score += full_text.count(kw)
        if score > 0:
            scores[cat['name']] = score
    if not scores:
        return '未分类'
    return max(scores, key=scores.get)

# ========== 4. 安全移动（含日志） ==========
def safe_move(src, dst_dir, log_file=None):
    dst_dir.mkdir(parents=True, exist_ok=True)
    dst = dst_dir / src.name
    base = src.stem
    ext = src.suffix
    counter = 1
    while dst.exists():
        dst = dst_dir / f'{base}_{counter}{ext}'
        counter += 1
    shutil.move(str(src), str(dst))
    if log_file:
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f'{src}|{dst}\n')
    return dst

# ========== 5. 主流程 ==========
def main():
    if not Path('categories.yaml').exists():
        print('错误：未找到 categories.yaml 配置文件，请将它与程序放在同一文件夹下。')
        input('按回车键退出...')
        return
    categories = load_categories()

    desktop = Path.home() / 'Desktop'
    archive_root = desktop / f'桌面整理_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
    log_path = archive_root / 'undo_log.txt'

    files = [
        f for f in desktop.iterdir()
        if f.is_file() and not f.name.startswith('.') and f.suffix.lower() not in ('.lnk', '.ini')
    ]
    if not files:
        print('桌面空空如也，没有需要整理的文件。')
        input('按回车键退出...')
        return

    print(f'找到 {len(files)} 个文件：')
    for f in files:
        print(f'  - {f.name}')
    print(f'\n整理后文件将移动到：{archive_root}')

    # 确认环节
    confirm = input('\n按回车键开始整理，输入 n 取消... ')
    if confirm.lower() == 'n':
        print('已取消整理。')
        input('按回车键退出...')
        return

    print('\n开始整理...\n')
    for file in files:
        print(f'处理: {file.name}')
        text = extract_text(file)
        if not text.strip():
            category = '未分类（无法读取内容）'
        else:
            category = classify_document(text, categories)
        try:
            dest = safe_move(file, archive_root / category, log_file=log_path)
            print(f'  ✅ 归类至: {category}')
        except Exception as e:
            print(f'  ❌ 移动失败: {e}')

    print('\n===== 整理完成 =====')
    print(f'归档位置: {archive_root}')
    print(f'如需撤销整理，请运行 undo.py，并输入日志文件路径：{log_path}')
    input('\n按回车键退出...')

if __name__ == '__main__':
    main()
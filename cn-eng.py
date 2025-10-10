from docx import Document
import pandas as pd
import os
import re

def read_word_content(docx_path):
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"Word文件不存在：{docx_path}")
    
    doc = Document(docx_path)
    title = ""
    paragraphs = []
    
    for para in doc.paragraphs:
        text = para.text.strip()
        
        # 跳过空段落
        if not text:
            continue
            
        # 提取标题（第一个非分页标记的段落）
        if not title and text != '@@page' and text != '@@pe':
            title = text
            
        # 添加所有段落，包括分页标记（在分页逻辑中处理）
        cleaned_text = re.sub(r'\s+', ' ', text)
        paragraphs.append(cleaned_text)
    
    if not paragraphs:
        raise ValueError("Word文档无有效内容")
    return title, paragraphs

def read_excel_words(xlsx_path):
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"Excel文件不存在：{xlsx_path}")
    
    df = pd.read_excel(xlsx_path, dtype=str, engine='openpyxl')
    word_dict = {}
    
    for _, row in df.iterrows():
        word = str(row.get('单词', '')).strip()
        if word:
            word_lower = word.lower()
            word_dict[word_lower] = {
                'original': word,
                'phonetic': str(row.get('音标', '')).strip(),
                'meaning': str(row.get('意思', '')).strip(),
                'synonym': str(row.get('同义词', '')).strip() or '无'
            }
    return word_dict

def split_into_pages(paragraphs, word_dict, min_words=50, max_words=70):
    pages = []
    current_page = []
    current_count = 0
    
    for para in paragraphs:
        # 检查是否是分页标记
        if para == '@@page':
            # 强制分页：如果当前页有内容，先保存当前页
            if current_page:
                page_text = ' '.join(current_page)
                page_words = extract_page_words(page_text, word_dict)
                pages.append((current_page.copy(), page_words))
                current_page = []
                current_count = 0
            # 开始新页（空页，分页标记本身不加入内容）
            continue
        elif para == '@@pe':
            # 结束标记，暂时忽略或用于其他逻辑
            continue
        
        # 计算段落单词数（只计算英文单词）
        words = re.findall(r"[a-zA-Z0-9']+", para)
        para_count = len([w for w in words if re.search(r'[a-zA-Z]', w)])
        
        # 自动分页逻辑：如果超过最大单词数且当前页有内容，则分页
        if current_count + para_count > max_words and current_page:
            page_text = ' '.join(current_page)
            page_words = extract_page_words(page_text, word_dict)
            pages.append((current_page.copy(), page_words))
            
            current_page = [para]
            current_count = para_count
        else:
            current_page.append(para)
            current_count += para_count
    
    # 处理最后一页
    if current_page:
        page_text = ' '.join(current_page)
        page_words = extract_page_words(page_text, word_dict)
        pages.append((current_page, page_words))
    
    return pages

def extract_page_words(page_text, word_dict):
    page_words = []
    used_words = set()
    
    words = re.findall(r"[a-zA-Z0-9']+", page_text)
    for w in words:
        clean_w = re.sub(r'[^a-zA-Z0-9]', '', w).lower()
        if clean_w in word_dict and clean_w not in used_words:
            page_words.append(word_dict[clean_w])
            used_words.add(clean_w)
            if len(page_words) >= 6:
                break
    
    return page_words

def generate_a4_html(title, paragraphs, word_dict, output, font_scale=1.5, right_font_scale=1.8):
    font_scale = max(0.8, min(2.0, font_scale))
    right_font_scale = max(1.0, min(2.5, right_font_scale))
    pages = split_into_pages(paragraphs, word_dict)
    pages_html = ""
    
    for i, (page_paras, page_words) in enumerate(pages, 1):
        left_html = ""
        for idx, para in enumerate(page_paras):
            highlighted = para
            for w in page_words:
                highlighted = re.sub(
                    rf"\b{re.escape(w['original'])}\b",
                    f'<span class="highlight">{w["original"]}</span>',
                    highlighted,
                    flags=re.IGNORECASE
                )
            left_html += f"<p>{highlighted}</p>"
            if idx < len(page_paras) - 1:
                left_html += '<p class="spacing"></p>'
        
        right_html = ""
        for w in page_words:
            right_html += f'''
            <div class="word-card">
                <div class="word">{w['original']}</div>
                <div class="phonetic">[{w['phonetic']}]</div>
                <div class="meaning">意思：{w['meaning']}</div>
                <div class="synonym">同义词：{w['synonym']}</div>
            </div>
            '''
        if not right_html:
            right_html = '<p class="no-words">无匹配单词</p>'
        
        pages_html += f'''
        <div class="page" data-page="{i}" style="--scale:{font_scale}; --right-scale:{right_font_scale}; --card-gap:0.8;">
            <div class="header">
                <h1>{title}</h1>
                <div class="page-num">第{i}/{len(pages)}页</div>
            </div>
            <div class="content">
                <div class="left">{left_html}</div>
                <div class="right">{right_html}</div>
            </div>
        </div>
        '''
    
    html = f'''<!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>{title}</title>
        <style>
            .page {{
                width: 210mm;
                height: 297mm;
                margin: 20px auto;
                padding: 15mm;
                background: white;
                box-shadow: 0 2px 5px #ccc;
                overflow: hidden;
                --scale: 1.5;
                --right-scale: 1.8;
                --card-gap: 0.8;
                --min-scale: 0.7;
                --min-card-gap: 0.3;
            }}
            
            .header {{
                display: flex;
                justify-content: space-between;
                margin-bottom: 10mm;
                padding-bottom: 5mm;
                border-bottom: 1px solid #eee;
            }}
            .header h1 {{
                font-size: calc(16pt * var(--scale) * 0.8);
                margin: 0;
            }}
            .page-num {{
                font-size: calc(10pt * var(--scale) * 0.7);
                color: #666;
            }}
            
            .content {{
                display: flex;
                gap: 10mm;
                height: calc(297mm - 80mm);
                position: relative;
            }}
            
            .left {{
                flex: 5;
                font-size: calc(12pt * var(--scale));
                line-height: 1.8;
                overflow: hidden;
                position: relative;
            }}
            .left p {{
                margin: 0 0 calc(5mm * var(--scale)) 0;
                text-indent: 2em;
                word-wrap: break-word;
            }}
            .spacing {{
                height: calc(8mm * var(--scale));
            }}
            .highlight {{
                background: #fff9c4;
                padding: 0 2px;
                font-weight: bold;
            }}
            
            .right {{
                flex: 4;
                padding-left: 5mm;
                border-left: 1px dashed #eee;
                overflow: hidden;
                position: relative;
            }}
            .word-card {{
                margin-bottom: calc(10mm * var(--right-scale) * var(--card-gap));
                padding: calc(5mm * var(--right-scale) * 0.8);
                background: #f9f9f9;
                border-radius: 4px;
                word-wrap: break-word;
            }}
            .word {{
                font-weight: bold;
                font-size: calc(14pt * var(--right-scale));
            }}
            .phonetic {{
                color: #0288d1;
                font-size: calc(12pt * var(--right-scale));
                margin: calc(2mm * var(--right-scale) * var(--card-gap)) 0;
                font-style: italic;
            }}
            .meaning {{
                color: #d32f2f;
                font-size: calc(13pt * var(--right-scale));
                margin: calc(2mm * var(--right-scale) * var(--card-gap)) 0;
            }}
            .synonym {{
                color: #7b1fa2;
                font-size: calc(13pt * var(--right-scale));
                margin: calc(2mm * var(--right-scale) * var(--card-gap)) 0;
            }}
            .no-words {{
                color: #999;
                text-align: center;
                margin-top: 20mm;
                font-size: calc(12pt * var(--right-scale));
            }}
            
            @media print {{
                .page {{
                    width: 100%;
                    height: 100%;
                    margin: 0;
                    padding: 0;
                    box-shadow: none;
                    page-break-after: always;
                }}
            }}
        </style>
    </head>
    <body>
        {pages_html}
        
        <script>
            window.addEventListener('load', function() {{
                const pages = document.querySelectorAll('.page');
                pages.forEach(page => adjustPageLayout(page));
            }});
            
            function adjustPageLayout(page) {{
                const rightSection = page.querySelector('.right');
                const leftSection = page.querySelector('.left');
                
                adjustRightSection(page, rightSection);
                adjustLeftSection(page, leftSection);
            }}
            
            function adjustRightSection(page, rightSection) {{
                let cardGap = parseFloat(getComputedStyle(page).getPropertyValue('--card-gap'));
                let rightScale = parseFloat(getComputedStyle(page).getPropertyValue('--right-scale'));
                const minCardGap = parseFloat(getComputedStyle(page).getPropertyValue('--min-card-gap'));
                const minScale = parseFloat(getComputedStyle(page).getPropertyValue('--min-scale'));
                const gapStep = 0.05;
                const scaleStep = 0.05;
                
                while (isOverflowing(rightSection) && cardGap > minCardGap) {{
                    cardGap = Math.max(cardGap - gapStep, minCardGap);
                    page.style.setProperty('--card-gap', cardGap.toFixed(2));
                    forceRepaint(rightSection);
                }}
                
                while (isOverflowing(rightSection) && rightScale > minScale) {{
                    rightScale = Math.max(rightScale - scaleStep, minScale);
                    page.style.setProperty('--right-scale', rightScale.toFixed(2));
                    forceRepaint(rightSection);
                }}
            }}
            
            function adjustLeftSection(page, leftSection) {{
                let leftScale = parseFloat(getComputedStyle(page).getPropertyValue('--scale'));
                const minScale = parseFloat(getComputedStyle(page).getPropertyValue('--min-scale'));
                const scaleStep = 0.05;
                
                while (isOverflowing(leftSection) && leftScale > minScale) {{
                    leftScale = Math.max(leftScale - scaleStep, minScale);
                    page.style.setProperty('--scale', leftScale.toFixed(2));
                    forceRepaint(leftSection);
                }}
            }}
            
            function isOverflowing(element) {{
                return element.scrollHeight > element.offsetHeight || 
                       element.scrollWidth > element.offsetWidth;
            }}
            
            function forceRepaint(element) {{
                element.offsetHeight;
            }}
        </script>
    </body>
    </html>'''
    
    with open(output, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"生成成功：{os.path.abspath(output)}")
    print(f"初始配置：左侧缩放{font_scale}倍，右侧缩放{right_font_scale}倍，右侧卡片间距系数0.8")

def main():
    docx_path = "666.docx"
    excel_path = "666.xlsx"
    output_path = "output.html"
    
    try:
        title, paras = read_word_content(docx_path)
        words = read_excel_words(excel_path)
        generate_a4_html(title, paras, words, output_path, font_scale=1.5, right_font_scale=1.8)
    except Exception as e:
        print(f"错误：{str(e)}")

if __name__ == "__main__":
    main()
from bs4 import BeautifulSoup
import re

def extract_word_pages(html_file_path="output.html"):
    """
    提取同级output.html中右侧含单词内容的页码（默认路径为同级output.html）
    :param html_file_path: HTML文件路径（默认同级output.html）
    :return: 右侧含单词的页码列表（整数类型）
    """
    # 读取同级的output.html文件
    try:
        with open(html_file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except FileNotFoundError:
        print(f"错误：未在同级目录找到 output.html 文件，请确认文件是否存在")
        return []
    except Exception as e:
        print(f"读取文件时发生错误：{str(e)}")
        return []

    # 解析HTML结构（核心逻辑不变）
    soup = BeautifulSoup(html_content, 'html.parser')
    page_elements = soup.find_all('div', class_='page')  # 匹配页面节点
    word_page_nums = []

    for page in page_elements:
        # 1. 提取页码（优先从data-page属性，其次从“第X/Y页”中匹配）
        data_page = page.get('data-page')
        if data_page and data_page.isdigit():
            page_num = int(data_page)
        else:
            page_num_element = page.find('div', class_='page-num')
            if not page_num_element:
                continue
            page_num_match = re.search(r'第(\d+)', page_num_element.get_text(strip=True))
            if not page_num_match:
                continue
            page_num = int(page_num_match.group(1))

        # 2. 检查右侧是否有单词卡片（.word-card）
        right_section = page.find('div', class_='right')
        if not right_section:
            continue
        if right_section.find_all('div', class_='word-card'):  # 存在单词卡片则记录页码
            word_page_nums.append(page_num)

    # 按页码排序
    word_page_nums.sort()
    return word_page_nums

def format_page_range(page_list):
    """将页码列表格式化为“连续用-、非连续用、”的字符串"""
    if not page_list:
        return "无右侧含单词的页面"
    
    formatted = []
    start = page_list[0]
    end = start

    for num in page_list[1:]:
        if num == end + 1:
            end = num
        else:
            formatted.append(f"{start}-{end}" if start != end else str(start))
            start = num
            end = num
    # 处理最后一段
    formatted.append(f"{start}-{end}" if start != end else str(start))
    
    return "、".join(formatted)

# 主执行逻辑（直接运行即可）
if __name__ == "__main__":
    print("正在分析同级目录的 output.html 文件...")
    word_pages = extract_word_pages()  # 默认读取同级output.html
    formatted_result = format_page_range(word_pages)
    
    # 输出结果
    print("=" * 60)
    print("分析完成！")
    print(f"右侧含单词内容的页码：{formatted_result}")
    print("=" * 60)
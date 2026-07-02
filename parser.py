"""
题库解析模块 v2 - 增强版
"""
import re
from docx import Document

def parse_question_bank(filepath):
    doc = Document(filepath)
    questions = []
    current_type = None       # 'single' / 'multiple' / 'judge'
    current_chapter = 1
    current_question = None
    in_question = False
    collected_options = []

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            if in_question and current_question and current_type in ('single', 'multiple'):
                _save(questions, current_question, current_type, current_chapter, collected_options)
                current_question = None
                in_question = False
                collected_options = []
            continue

        # === 新章节检测：一、单选 / 一、单项选择题 ===
        if re.match(r'^一[、.]\s*(?:单项选择题|单选题|单选)$', text):
            current_type = 'single'
            if questions:
                current_chapter += 1
            continue

        # === 题型切换 ===
        if re.match(r'^[二二][、.]\s*(?:多项选择题|多选题|多选)$', text):
            current_type = 'multiple'
            continue
        if re.match(r'^[三三][、.]\s*(?:判断题|判断)$', text):
            current_type = 'judge'
            continue
        # 不带序号的题型标题
        if text in ('多项选择题', '多选题', '多选'):
            current_type = 'multiple'
            continue
        if text in ('判断题', '判断'):
            current_type = 'judge'
            continue

        # === 题目开头 ===
        q_match = re.match(r'^(\d+)[\.\s、\)]+(.+)', text)
        if q_match:
            if in_question and current_question and current_type in ('single', 'multiple'):
                _save(questions, current_question, current_type, current_chapter, collected_options)
                collected_options = []

            q_num = int(q_match.group(1))
            q_text = q_match.group(2)

            if current_type == 'judge':
                questions.append({
                    'type': 'judge', 'number': q_num, 'chapter': current_chapter,
                    'question': q_text, 'answer': _extract_answer(q_text),
                    'options': [], 'score': 1
                })
                current_question = None
                in_question = False
                continue

            current_question = q_text
            in_question = True
            continue

        # === 选项行 ===
        opt_match = re.match(r'^([A-E])[\.\s、]+(.+)', text)
        if in_question and current_type in ('single', 'multiple') and opt_match:
            collected_options.append({
                'label': opt_match.group(1),
                'text': opt_match.group(2).strip()
            })
            continue

        # === 题干续行 ===
        if in_question and current_question:
            current_question += text

    # 最后一题
    if in_question and current_question and current_type in ('single', 'multiple'):
        _save(questions, current_question, current_type, current_chapter, collected_options)

    return questions


def _extract_answer(text):
    """从题目文本提取答案"""
    m = re.search(r'[（(]\s*([A-D✓×√xX对錯错]+)\s*[）)]', text)
    if m:
        ans = m.group(1).strip().upper()
        ans = ans.replace('✓','√').replace('✗','×').replace('X','×')
        ans = ans.replace('对','√').replace('錯','×').replace('错','×')
        return ans
    return ''


def _extract_inline_options(text):
    """从题干中提取内联选项（选项和题目在同一行的情况）"""
    options = []
    # 匹配 A.xxx B.xxx C.xxx D.xxx 或 A、xxx B、xxx 格式
    pattern = re.compile(r'([A-E])[\.\s、]+(.+?)(?=[A-E][\.\s、]|$)', re.DOTALL)
    matches = pattern.findall(text)
    for label, content in matches:
        # 清理内容：去掉尾部多余空格和符号
        content = content.strip().rstrip('；;，,。.').strip()
        if content and len(content) < 200:  # 合理长度的选项
            options.append({'label': label, 'text': content})
    return options


def _save(questions, q_text, q_type, chapter, options):
    if q_type not in ('single', 'multiple'):
        return
    # 如果没有独立的选项行，尝试从题干中提取内联选项
    if not options:
        options = _extract_inline_options(q_text)
    answer = _extract_answer(q_text)
    # 去掉答案标记
    clean = re.sub(r'[（(]\s*[A-D✓×√xX对錯错]+\s*[）)]', '（  ）', q_text).strip()
    num_match = re.match(r'^(\d+)', clean)
    number = int(num_match.group(1)) if num_match else 0
    clean = re.sub(r'^\d+[\.\s、\)]+', '', clean)

    questions.append({
        'type': q_type, 'number': number, 'chapter': chapter,
        'question': clean, 'answer': answer, 'options': options, 'score': 1
    })


def get_statistics(questions):
    from collections import defaultdict
    stats = {'single':0,'multiple':0,'judge':0,'total':len(questions)}
    ch_stats = defaultdict(lambda:{'single':0,'multiple':0,'judge':0})
    for q in questions:
        stats[q['type']] = stats.get(q['type'], 0) + 1
        ch_stats[q['chapter']][q['type']] += 1
    stats['chapters'] = sorted(ch_stats.keys())
    stats['chapter_detail'] = dict(ch_stats)
    return stats

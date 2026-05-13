#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""初中化学学案二次备课批注工具 - 使用大模型生成批注"""

import sys
import os
import io
import shutil
import zipfile
import re
import json
import requests
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

SKILL_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "agent-skills包", "docx"))
sys.path.insert(0, SKILL_ROOT)
from scripts.document import Document


def read_docx_content(unpacked_dir):
    """读取docx文档内容，返回段落列表和位置信息"""
    import defusedxml.minidom
    
    dom = defusedxml.minidom.parse(str(Path(unpacked_dir) / "word" / "document.xml"))
    
    paragraphs = []
    for idx, para in enumerate(dom.getElementsByTagName("w:p")):
        texts = []
        for r in para.getElementsByTagName("w:r"):
            for t in r.getElementsByTagName("w:t"):
                if t.firstChild:
                    texts.append(t.firstChild.nodeValue)
        
        if texts:
            full_text = "".join(texts).strip()
            if full_text and len(full_text) > 5:
                paragraphs.append({
                    "index": idx,
                    "text": full_text,
                    "node": para
                })
    
    # 处理表格
    for tbl in dom.getElementsByTagName("w:tbl"):
        for row in tbl.getElementsByTagName("w:tr"):
            for cell in row.getElementsByTagName("w:tc"):
                cell_paras = cell.getElementsByTagName("w:p")
                for p in cell_paras:
                    texts = []
                    for r in p.getElementsByTagName("w:r"):
                        for t in r.getElementsByTagName("w:t"):
                            if t.firstChild:
                                texts.append(t.firstChild.nodeValue)
                    if texts:
                        full_text = "".join(texts).strip()
                        if full_text and len(full_text) > 5:
                            paragraphs.append({
                                "index": -1,
                                "text": full_text,
                                "node": p,
                                "in_table": True
                            })
    
    return paragraphs


def generate_all_from_api(paragraphs, segments_text="", api_key=None):
    """一次性调用API生成批注、教学反思和时间分配"""
    if not api_key:
        api_key = os.environ.get("DEEPSEEK_API_KEY")
    
    if not api_key:
        # 返回模拟数据
        annotations = generate_mock_annotations(paragraphs)
        reflection = "本节课基本达成教学目标，学生通过练习对所学知识有了更深入的理解。教学重点通过课堂讲解和例题分析得到强化，部分难点仍需加强练习。在今后的教学中，应注重知识的系统性和学生能力的培养，进一步提高教学效果。"
        time_allocation = None
        return annotations, reflection, time_allocation
    
    # 准备文档内容
    content_parts = []
    for p in paragraphs:
        content_parts.append(f"[段落{p['index']}] {p['text']}")
    document_text = "\n\n".join(content_parts)
    
    prompt = f"""你是一位有经验的初中化学教师，正在审阅一份教案。

请一次性生成以下三部分内容：

1. 批注（15-20条，每条80字以上，均匀分布在文档开头、中间、结尾）
2. 教学反思（约200字，包括：教学目标达成情况、教学重难点处理、学生掌握情况、改进措施）
3. 时间分配（总时长45分钟，开始环节短，中间环节长，结尾环节短，格式如：5,15,10,10,5）

批注要求：
- 像同事交流，语气自然
- 不用【】符号，不用"你"字
- 用"这段"、"这个"等

教案内容：
{document_text}

{segments_text}

请按以下JSON格式输出（只输出JSON，不要其他内容）：
{{
  "annotations": [
    {{"paragraph_index": 编号, "comment": "批注内容"}},
    ...
  ],
  "reflection": "教学反思内容",
  "time_allocation": "5,15,10,10,5"
}}"""

    try:
        print("调用DeepSeek API生成批注、教学反思、时间分配...")
        
        response = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            },
            json={
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": "你是一位有经验的初中化学教师，擅长给出实用、有针对性的教学建议。"},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 4000
            },
            timeout=120
        )
        
        result = response.json()
        
        if "error" in result:
            print(f"API错误: {result['error']['message']}")
            annotations = generate_mock_annotations(paragraphs)
            reflection = "本节课基本达成教学目标，学生通过练习对所学知识有了更深入的理解。教学重点通过课堂讲解和例题分析得到强化，部分难点仍需加强练习。在今后的教学中，应注重知识的系统性和学生能力的培养，进一步提高教学效果。"
            return annotations, reflection, None
        
        result_text = result["choices"][0]["message"]["content"].strip()
        
        # 提取JSON
        if "```json" in result_text:
            result_text = result_text.split("```json")[1].split("```")[0]
        elif "```" in result_text:
            result_text = result_text.split("```")[1].split("```")[0]
        
        # 解析JSON
        import json
        data = json.loads(result_text)
        
        annotations = data.get("annotations", [])
        reflection = data.get("reflection", "")
        time_str = data.get("time_allocation", "")
        
        # 解析时间分配
        time_allocation = []
        for part in time_str.replace(',', ' ').split():
            try:
                t = int(part.strip())
                if 0 < t <= 60:
                    time_allocation.append(t)
            except:
                pass
        
        print(f"API返回: {len(annotations)}条批注, 反思长度{len(reflection)}字, 时间分配{time_allocation}")
        return annotations, reflection, time_allocation
        
    except Exception as e:
        print(f"API调用失败: {e}")
        annotations = generate_mock_annotations(paragraphs)
        reflection = "本节课基本达成教学目标，学生通过练习对所学知识有了更深入的理解。教学重点通过课堂讲解和例题分析得到强化，部分难点仍需加强练习。在今后的教学中，应注重知识的系统性和学生能力的培养，进一步提高教学效果。"
        return annotations, reflection, None
    
    # 准备文档内容
    total = len(paragraphs)
    third = total // 3
    
    content_parts = []
    for p in paragraphs:
        content_parts.append(f"[段落{p['index']}] {p['text']}")
    
    document_text = "\n\n".join(content_parts)
    
    prompt = f"""你是一位有经验的初中化学教师，正在审阅一份教案（共{len(paragraphs)}个有效段落）。
请为教案的不同部分添加批注，确保批注均匀分布。

批注要求：
1. 生成15-20条批注，每条80字以上
2. 批注必须均匀分布在文档各处：
   - 开头部分（段落{paragraphs[0]['index']}到段落{paragraphs[third-1]['index'] if third < len(paragraphs) else paragraphs[-1]['index']}）：选4-7段
   - 中间部分（段落{paragraphs[third]['index'] if third < len(paragraphs) else paragraphs[-1]['index']}到段落{paragraphs[third*2-1]['index'] if third*2 <= len(paragraphs) else paragraphs[-1]['index']}）：选5-8段
   - 结尾部分（段落{paragraphs[third*2]['index'] if third*2 < len(paragraphs) else paragraphs[-1]['index']}到段落{paragraphs[-1]['index']}）：选4-7段
3. 像同事交流一样，语气自然，不要用"建议"，直接说
4. 批注要针对这段内容本身，而不是泛泛而谈
5. 不要用【】等符号开头
6. 不要出现"你"字，用"这段"、"这个"等

输出格式（JSON数组，每个元素包含paragraph_index和comment）：
[
  {{"paragraph_index": 段落编号, "comment": "批注内容"}},
  {{"paragraph_index": 段落编号, "comment": "批注内容"}}
]

教案内容：
{document_text}

请生成批注（只输出JSON，不要其他内容）："""

    try:
        print("调用DeepSeek API生成批注...")
        
        response = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            },
            json={
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": "你是一位有经验的初中化学教师，擅长给出实用、有针对性的教学建议。"},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.7,
                "max_tokens": 4000
            },
            timeout=120
        )
        
        result = response.json()
        
        if "error" in result:
            print(f"API错误: {result['error']['message']}")
            return generate_mock_annotations(paragraphs)
        
        result_text = result["choices"][0]["message"]["content"].strip()
        
        # 提取JSON
        if "```json" in result_text:
            result_text = result_text.split("```json")[1].split("```")[0]
        elif "```" in result_text:
            result_text = result_text.split("```")[1].split("```")[0]
        
        # 尝试解析JSON
        try:
            annotations = json.loads(result_text)
        except:
            # 尝试修复JSON
            result_text = re.sub(r'```json|```', '', result_text)
            annotations = json.loads(result_text)
        
        print(f"DeepSeek生成了 {len(annotations)} 条批注")
        return annotations
        
    except Exception as e:
        print(f"DeepSeek API调用失败: {e}")
        print("使用模拟数据...")
        return generate_mock_annotations(paragraphs)


def generate_mock_annotations(paragraphs):
    """生成模拟批注（当没有API key时使用）"""
    annotations = []
    
    for i, p in enumerate(paragraphs[:18]):
        text = p["text"]
        
        if "实验" in text or "操作" in text:
            comment = f"这个实验设计很好，能帮助学生直观理解。实验前要提醒学生观察重点，实验中注意规范操作，实验后引导学生分析现象、归纳结论。学生亲眼看到变化，理解会更深刻。"
        elif "讨论" in text or "探究" in text or "活动" in text:
            comment = f"这个活动设计能调动学生参与。小组讨论要注意明确分工和时间要求，每个人都要有任务。可以让学生代表展示讨论成果，这样参与度更高。"
        elif "例" in text or "解" in text or "计算" in text:
            comment = f"这道例题选得好，有代表性。讲评时不要只给答案，要分析思路和方法，让学生学会分析问题的方法。"
        elif "重点" in text or "难点" in text or "关键" in text:
            comment = f"这部分是教学重点，时间要留足。可以用多种方式帮助学生理解，比如举例、画图、对比等，让学生真正掌握而不是死记硬背。"
        elif "小结" in text or "总结" in text:
            comment = f"课堂小结能帮助学生梳理本节内容。可以让学生自己总结今天学了什么，教师补充完善，这样学生印象更深刻。"
        elif "作业" in text or "练习" in text:
            comment = f"作业要适量，基础题巩固知识，能力题培养应用。批改后及时反馈，针对典型错误集中讲解，这样效果更好。"
        elif "？" in text or "为什么" in text:
            comment = f"这个问题设计得不错，能引导学生思考。可以给学生一点时间想一想再回答，让学生自己得出答案印象更深刻。"
        else:
            comment = f"这段内容设计得不错。在实际教学中注意关注学生反应，灵活调整教学策略，让学生真正理解和掌握。"
        
        annotations.append({
            "paragraph_index": i,
            "comment": comment
        })
    
    return annotations


def add_annotations_to_doc(unpacked_dir, paragraphs, annotations, doc=None):
    """将批注添加到文档"""
    if doc is None:
        doc = Document(unpacked_dir, author="化学老师", initials="化")
    
    dom = doc["word/document.xml"].dom
    all_paragraphs = list(dom.getElementsByTagName("w:p"))
    
    index_to_node = {}
    for i, p in enumerate(all_paragraphs):
        index_to_node[i] = p
    
    added_count = 0
    
    for ann in annotations:
        para_idx = ann.get("paragraph_index", 0)
        
        if para_idx in index_to_node:
            node = index_to_node[para_idx]
            comment_text = ann.get("comment", "")
            
            if comment_text:
                try:
                    doc.add_comment(start=node, end=node, text=comment_text)
                    added_count += 1
                except Exception as e:
                    print(f"添加批注失败: {e}")
    
    print(f"成功添加 {added_count} 条批注")
    return doc


def generate_reflection(paragraphs, api_key=None):
    """调用API生成教学反思"""
    if not api_key:
        api_key = os.environ.get("DEEPSEEK_API_KEY")
    
    if not api_key:
        # 返回模拟数据
        return "本节课基本达成教学目标，学生通过练习对所学知识有了更深入的理解。教学重点通过课堂讲解和例题分析得到强化，部分难点仍需加强练习。在今后的教学中，应注重知识的系统性和学生能力的培养，进一步提高教学效果。"
    
    doc_summary = [p['text'][:100] for p in paragraphs[:20] if p['text'] and len(p['text']) > 10]
    
    prompt = f"""你是一位初中化学教师，请为以下教案生成约200字的教学反思。

教案内容：
{chr(10).join(doc_summary[:12])}

请从以下几个方面生成反思：
1. 教学目标的达成情况
2. 教学重难点的处理
3. 学生掌握情况
4. 改进措施

要求：约200字，语言自然流畅"""

    try:
        response = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
            json={"model": "deepseek-chat", "messages": [
                {"role": "system", "content": "你是一位经验丰富的初中化学教师"},
                {"role": "user", "content": prompt}
            ], "temperature": 0.7, "max_tokens": 500},
            timeout=60
        )
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        print(f"生成反思失败: {e}")
        return None


def add_reflection_to_table(unpacked_dir, reflection_text):
    """在表格的教学反思行添加反思内容"""
    import defusedxml.minidom
    
    doc_xml = Path(unpacked_dir) / "word" / "document.xml"
    dom = defusedxml.minidom.parse(str(doc_xml))
    
    tables = dom.getElementsByTagName("w:tbl")
    for tbl in tables:
        rows = tbl.getElementsByTagName("w:tr")
        for row in rows:
            cells = row.getElementsByTagName("w:tc")
            cell_texts = []
            for cell in cells:
                texts = []
                for p in cell.getElementsByTagName("w:p"):
                    for r in p.getElementsByTagName("w:r"):
                        for t in r.getElementsByTagName("w:t"):
                            if t.firstChild:
                                texts.append(t.firstChild.nodeValue)
                cell_texts.append(''.join(texts).strip())
            
            if any("教学反思" in t for t in cell_texts):
                if len(cells) >= 2:
                    target_cell = cells[1]
                    
                    # 清空该单元格内容：移除所有p元素
                    while target_cell.hasChildNodes():
                        target_cell.removeChild(target_cell.firstChild)
                    
                    # 添加新段落写入反思内容
                    new_para = dom.createElement("w:p")
                    new_r = dom.createElement("w:r")
                    new_t = dom.createElement("w:t")
                    new_t.appendChild(dom.createTextNode(reflection_text))
                    new_r.appendChild(new_t)
                    new_para.appendChild(new_r)
                    target_cell.appendChild(new_para)
                    
                    # 保存
                    with open(str(doc_xml), "w", encoding="utf-8") as f:
                        dom.writexml(f, indent="", encoding="utf-8")
                    
                    print("已在教学反思单元格填写反思内容")
                    return True
    
    print("未找到教学反思单元格")
    return False


def main():
    if len(sys.argv) < 2:
        print("用法: python annotate_doc.py <学案.docx> [输出文件.docx]")
        print("环境变量: DEEPSEEK_API_KEY (必填)")
        sys.exit(1)
    
    input_file = sys.argv[1]
    base_name = Path(input_file).stem
    from datetime import datetime
    current_time = datetime.now().strftime("%Y.%m.%d")
    output_file = sys.argv[2] if len(sys.argv) > 2 else str(Path(input_file).parent / f"{base_name}_二次备课_{current_time}.docx")
    unpack_dir = input_file.replace(".docx", "_unpacked")
    
    print(f"处理: {input_file}")
    
    # 解包
    print("解包文档...")
    shutil.rmtree(unpack_dir, ignore_errors=True)
    zipfile.ZipFile(input_file).extractall(unpack_dir)
    
    # 第一步：读取文档内容
    print("读取文档内容...")
    paragraphs = read_docx_content(unpack_dir)
    print(f"  文档段落数: {len(paragraphs)}")
    print(f"  总字数: {sum(len(p['text']) for p in paragraphs)}")
    
    # 第二步：一次性调用API生成批注、教学反思、时间分配
    print("生成批注、教学反思、时间分配...")
    annotations, reflection, time_allocation = generate_all_from_api(paragraphs)
    
    # 第三步：添加批注到文档
    print("添加批注到文档...")
    doc = add_annotations_to_doc(unpack_dir, paragraphs, annotations)
    
    # 第四步：保存文档
    print("保存文档...")
    doc.save(validate=False)
    with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in Path(unpack_dir).rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(unpack_dir))
    
    # 第五步：添加教学反思到表格（在打包后操作）
    if reflection:
        print("添加教学反思...")
        add_reflection_to_table_after_zip(output_file, reflection)
    
    # 第六步：添加时间分配
    if time_allocation:
        add_time_allocation_with_values(output_file, time_allocation)
    else:
        add_time_allocation(output_file)
    
    shutil.rmtree(unpack_dir, ignore_errors=True)
    print(f"\n完成: {output_file}")
    print("在Word中打开文档，点击审阅->批注查看")
    print("教学反思已填写到表格中")


def add_reflection_to_table_after_zip(output_file, reflection_text):
    """在打包后的docx文件中添加教学反思"""
    import defusedxml.minidom
    
    unpack_dir = output_file.replace(".docx", "_reflect_unpacked")
    
    # 解包
    shutil.rmtree(unpack_dir, ignore_errors=True)
    zipfile.ZipFile(output_file).extractall(unpack_dir)
    
    doc_xml = Path(unpack_dir) / "word" / "document.xml"
    dom = defusedxml.minidom.parse(str(doc_xml))
    
    tables = dom.getElementsByTagName("w:tbl")
    for tbl in tables:
        rows = tbl.getElementsByTagName("w:tr")
        for row in rows:
            cells = row.getElementsByTagName("w:tc")
            cell_texts = []
            for cell in cells:
                texts = []
                for p in cell.getElementsByTagName("w:p"):
                    for r in p.getElementsByTagName("w:r"):
                        for t in r.getElementsByTagName("w:t"):
                            if t.firstChild:
                                texts.append(t.firstChild.nodeValue)
                cell_texts.append(''.join(texts).strip())
            
            if any("教学反思" in t for t in cell_texts):
                if len(cells) >= 2:
                    target_cell = cells[1]
                    
                    # 只清空文本内容，保留段落结构
                    paras = target_cell.getElementsByTagName("w:p")
                    for para in paras:
                        runs = para.getElementsByTagName("w:r")
                        for run in runs:
                            text_nodes = run.getElementsByTagName("w:t")
                            for t in text_nodes:
                                # 清空文本内容
                                if t.firstChild:
                                    t.firstChild.nodeValue = ""
                    
                    # 在第一个段落中添加文本
                    if paras:
                        first_para = paras[0]
                        runs = first_para.getElementsByTagName("w:r")
                        if runs:
                            first_run = runs[0]
                            text_nodes = first_run.getElementsByTagName("w:t")
                            if text_nodes:
                                # 追加文本到现有文本节点
                                text_nodes[0].appendChild(dom.createTextNode(reflection_text))
                            else:
                                new_t = dom.createElement("w:t")
                                new_t.appendChild(dom.createTextNode(reflection_text))
                                first_run.appendChild(new_t)
                        else:
                            new_r = dom.createElement("w:r")
                            new_t = dom.createElement("w:t")
                            new_t.appendChild(dom.createTextNode(reflection_text))
                            new_r.appendChild(new_t)
                            first_para.appendChild(new_r)
                    
                    # 保存
                    with open(str(doc_xml), "w", encoding="utf-8") as f:
                        dom.writexml(f, indent="", encoding="utf-8")
                    
                    # 重新打包
                    with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zf:
                        for f in Path(unpack_dir).rglob("*"):
                            if f.is_file():
                                zf.write(f, f.relative_to(unpack_dir))
                    
                    shutil.rmtree(unpack_dir, ignore_errors=True)
                    print("已在教学反思单元格填写反思内容")
                    return
    
    shutil.rmtree(unpack_dir, ignore_errors=True)
    print("未找到教学反思单元格")


def get_time_allocation_from_api(segments, api_key=None):
    """调用API生成合理的时间分配"""
    if not api_key:
        api_key = os.environ.get("DEEPSEEK_API_KEY")
    
    if not api_key:
        return None
    
    segment_names = [s['text'][:30] for s in segments]
    
    prompt = f"""你是一位初中化学教师，请为以下{len(segment_names)}个教学环节分配时间。

环节列表：
{chr(10).join([f"{i+1}. {name}" for i, name in enumerate(segment_names)])}

要求：
1. 总时间45分钟
2. 开始环节（导入、情境等）时间短一些
3. 中间环节（重点讲解、练习、活动等）时间长一些
4. 结尾环节（总结、检测等）时间短一些
5. 给出每个环节的时间（分钟）

请只输出各环节时间，用逗号分隔，如：5,15,15,5,5"""

    try:
        response = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
            json={"model": "deepseek-chat", "messages": [
                {"role": "system", "content": "你是一位经验丰富的初中化学教师"},
                {"role": "user", "content": prompt}
            ], "temperature": 0.3, "max_tokens": 100},
            timeout=30
        )
        result = response.json()["choices"][0]["message"]["content"].strip()
        
        # 解析结果
        times = []
        for part in result.replace(',', ' ').split():
            try:
                t = int(part.strip())
                if 0 < t <= 60:
                    times.append(t)
            except:
                pass
        
        if len(times) == len(segments) and sum(times) == 45:
            return times
        
        # 如果解析失败，使用默认分配
        return None
    except:
        return None


def add_time_allocation(output_file):
    """为教学环节添加时间分配标记"""
    import defusedxml.minidom
    
    unpack_dir = output_file.replace(".docx", "_time_unpacked")
    
    # 解包
    shutil.rmtree(unpack_dir, ignore_errors=True)
    zipfile.ZipFile(output_file).extractall(unpack_dir)
    
    doc_xml = Path(unpack_dir) / "word" / "document.xml"
    dom = defusedxml.minidom.parse(str(doc_xml))
    
    # 查找所有教学环节
    tables = dom.getElementsByTagName("w:tbl")
    segments = []
    
    for tbl in tables:
        rows = tbl.getElementsByTagName("w:tr")
        for row_idx, row in enumerate(rows):
            cells = row.getElementsByTagName("w:tc")
            if len(cells) > 0:
                cell = cells[0]
                texts = []
                for p in cell.getElementsByTagName("w:p"):
                    for r in p.getElementsByTagName("w:r"):
                        for t in r.getElementsByTagName("w:t"):
                            if t.firstChild:
                                texts.append(t.firstChild.nodeValue)
                full_text = ''.join(texts).strip()
                # 匹配：环节X开头，且有实际内容（排除"环节"表头）
                if full_text.startswith("环节") and len(full_text) > 2:
                    segments.append({'row_idx': row_idx, 'cell': cell, 'text': full_text})
    
    if not segments:
        print("未找到教学环节，跳过时间分配")
        shutil.rmtree(unpack_dir, ignore_errors=True)
        return
    
    # 检查是否已有时间标记
    has_time = any("分钟" in s['text'] or "min" in s['text'].lower() or "Min" in s['text'] for s in segments)
    
    if has_time:
        print("教学环节已有时间分配，跳过")
        shutil.rmtree(unpack_dir, ignore_errors=True)
        return
    
    # 尝试调用API获取合理的时间分配
    times = get_time_allocation_from_api(segments)
    
    if not times:
        # 使用默认分配：开始和结尾短，中间长
        n = len(segments)
        if n <= 0:
            shutil.rmtree(unpack_dir, ignore_errors=True)
            return
        
        # 中间环节时间长，两端短
        if n == 1:
            times = [45]
        elif n == 2:
            times = [15, 30]
        elif n == 3:
            times = [10, 25, 10]
        elif n == 4:
            times = [8, 20, 12, 5]
        elif n == 5:
            times = [8, 18, 12, 5, 2]
        elif n == 6:
            times = [5, 15, 12, 8, 3, 2]
        else:
            # 均匀分配
            base_time = 45 // n
            times = []
            for i in range(n):
                if i < 45 % n:
                    times.append(base_time + 1)
                else:
                    times.append(base_time)
        
        print(f"使用默认时间分配: {times}")
    else:
        print(f"API生成时间分配: {times}")
    
    # 在每个环节文本后添加时间标记（红色）
    for seg, time in zip(segments, times):
        cell = seg['cell']
        paras = cell.getElementsByTagName("w:p")
        
        if paras:
            # 找到包含"环节"的段落
            for para in paras:
                texts = []
                for r in para.getElementsByTagName("w:r"):
                    for t in r.getElementsByTagName("w:t"):
                        if t.firstChild:
                            texts.append(t.firstChild.nodeValue)
                full_text = ''.join(texts).strip()
                
                if "环节" in full_text:
                    # 添加时间标记：创建带颜色的run
                    # 添加空格
                    space_run = dom.createElement("w:r")
                    space_t = dom.createElement("w:t")
                    space_t.appendChild(dom.createTextNode(" "))
                    space_run.appendChild(space_t)
                    para.appendChild(space_run)
                    
                    # 添加时间run（红色）
                    time_run = dom.createElement("w:r")
                    
                    # 添加颜色属性（红色）
                    rPr = dom.createElement("w:rPr")
                    color = dom.createElement("w:color")
                    color.setAttribute("w:val", "FF0000")
                    rPr.appendChild(color)
                    time_run.appendChild(rPr)
                    
                    time_t = dom.createElement("w:t")
                    time_t.appendChild(dom.createTextNode(f"{time}min"))
                    time_run.appendChild(time_t)
                    
                    para.appendChild(time_run)
                    break
    
    # 保存
    with open(str(doc_xml), "w", encoding="utf-8") as f:
        dom.writexml(f, indent="", encoding="utf-8")
    
    # 重新打包
    with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in Path(unpack_dir).rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(unpack_dir))
    
    shutil.rmtree(unpack_dir, ignore_errors=True)
    print(f"已添加时间分配（{len(segments)}个环节，总计{sum(times)}分钟）")


def add_time_allocation_with_values(output_file, time_allocation):
    """使用预分配的时间值为教学环节添加标记"""
    import defusedxml.minidom
    
    if not time_allocation:
        add_time_allocation(output_file)
        return
    
    unpack_dir = output_file.replace(".docx", "_time_unpacked")
    
    # 解包
    shutil.rmtree(unpack_dir, ignore_errors=True)
    zipfile.ZipFile(output_file).extractall(unpack_dir)
    
    doc_xml = Path(unpack_dir) / "word" / "document.xml"
    dom = defusedxml.minidom.parse(str(doc_xml))
    
    # 查找所有教学环节
    tables = dom.getElementsByTagName("w:tbl")
    segments = []
    
    for tbl in tables:
        rows = tbl.getElementsByTagName("w:tr")
        for row_idx, row in enumerate(rows):
            cells = row.getElementsByTagName("w:tc")
            if len(cells) > 0:
                cell = cells[0]
                texts = []
                for p in cell.getElementsByTagName("w:p"):
                    for r in p.getElementsByTagName("w:r"):
                        for t in r.getElementsByTagName("w:t"):
                            if t.firstChild:
                                texts.append(t.firstChild.nodeValue)
                full_text = ''.join(texts).strip()
                if full_text.startswith("环节") and len(full_text) > 2:
                    segments.append({'row_idx': row_idx, 'cell': cell, 'text': full_text})
    
    if not segments:
        print("未找到教学环节，跳过时间分配")
        shutil.rmtree(unpack_dir, ignore_errors=True)
        return
    
    # 检查是否已有时间标记
    has_time = any("分钟" in s['text'] or "min" in s['text'].lower() for s in segments)
    if has_time:
        print("教学环节已有时间分配，跳过")
        shutil.rmtree(unpack_dir, ignore_errors=True)
        return
    
    # 验证时间分配数量
    times = time_allocation[:len(segments)]
    if len(times) != len(segments) or sum(times) != 45:
        print(f"时间分配不匹配，使用默认分配")
        add_time_allocation(output_file)
        shutil.rmtree(unpack_dir, ignore_errors=True)
        return
    
    # 在每个环节文本后添加时间标记（红色）
    for seg, time in zip(segments, times):
        cell = seg['cell']
        paras = cell.getElementsByTagName("w:p")
        
        if paras:
            for para in paras:
                texts = []
                for r in para.getElementsByTagName("w:r"):
                    for t in r.getElementsByTagName("w:t"):
                        if t.firstChild:
                            texts.append(t.firstChild.nodeValue)
                full_text = ''.join(texts).strip()
                
                if "环节" in full_text:
                    # 添加空格
                    space_run = dom.createElement("w:r")
                    space_t = dom.createElement("w:t")
                    space_t.appendChild(dom.createTextNode(" "))
                    space_run.appendChild(space_t)
                    para.appendChild(space_run)
                    
                    # 添加时间run（红色）
                    time_run = dom.createElement("w:r")
                    rPr = dom.createElement("w:rPr")
                    color = dom.createElement("w:color")
                    color.setAttribute("w:val", "FF0000")
                    rPr.appendChild(color)
                    time_run.appendChild(rPr)
                    
                    time_t = dom.createElement("w:t")
                    time_t.appendChild(dom.createTextNode(f"{time}min"))
                    time_run.appendChild(time_t)
                    
                    para.appendChild(time_run)
                    break
    
    # 保存
    with open(str(doc_xml), "w", encoding="utf-8") as f:
        dom.writexml(f, indent="", encoding="utf-8")
    
    # 重新打包
    with zipfile.ZipFile(output_file, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in Path(unpack_dir).rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(unpack_dir))
    
    shutil.rmtree(unpack_dir, ignore_errors=True)
    print(f"API时间分配: {times}")


if __name__ == "__main__":
    main()

import gradio as gr
from docx import Document
import hanlp
from hanlp.pretrained.ner import MSRA_NER_BERT_BASE_ZH
import re
import io

hanlp_ner_model = hanlp.load(MSRA_NER_BERT_BASE_ZH)


def desensitize_docx(uploaded_files):
    desensitized_texts = []

    for uploaded_file in uploaded_files:
        # 读取上传的 .docx 文件
        doc = Document(io.BytesIO(uploaded_file))
        paragraphs = [para.text for para in doc.paragraphs]
        # 脱敏处理
        desensitized_paragraphs = [desensitize_with_hanlp(paragraph) for paragraph in paragraphs]
        desensitized_text = '\n'.join(desensitized_paragraphs)
        desensitized_texts.append(desensitized_text)

    return "\n\n---\n\n".join(desensitized_texts)  # 将多个文件的结果分隔开

def desensitize_with_hanlp(text):
    try:
        entities = hanlp_ner_model(text)
        desensitized_text = text
        for entity in reversed(entities):
            entity_text, label, start, end = entity
            if label in ['PERSON', 'LOCATION']:  # 根据需要调整实体类型
                unit = re.search(r'(局|公司|工程|省|市|县|区|街道|社区|小区|花园|苑)$', entity_text)
                if unit:
                    replacement = '**' + unit.group()
                    desensitized_text = desensitized_text[:start] + replacement + desensitized_text[end:]
                else:
                    desensitized_text = desensitized_text[:start] + '**' + desensitized_text[end:]
        return desensitized_text
    except Exception as e:
        return "Error in desensitizing: " + str(e)

# 创建 Gradio 应用
gr_interface = gr.Interface(
    title = "脱敏处理",
    fn=desensitize_docx,
    inputs=gr.File(type="binary", label="上传DOCX文件"),
    outputs="text"
)

# 运行应用
gr_interface.launch()

import sys
import re
import os
import random
import string
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QProgressBar, QTextEdit, QFileDialog
from PyQt5.QtCore import QThread, pyqtSignal, QObject
from docx import Document
import hanlp
from hanlp.pretrained.ner import MSRA_NER_BERT_BASE_ZH

hanlp_ner_model = hanlp.load(MSRA_NER_BERT_BASE_ZH)

def split_text(text, max_length):
    sentences = re.split(r'(?<=[。，”“！？\?!])', text)
    current_chunk = ""
    chunks = []

    for sentence in sentences:
        if len(current_chunk) + len(sentence) <= max_length:
            current_chunk += sentence
        else:
            if current_chunk:
                chunks.append(current_chunk)
            if len(sentence) > max_length:
                sub_chunks = [sentence[i:i+max_length] for i in range(0, len(sentence), max_length)]
                chunks.extend(sub_chunks)
            else:
                current_chunk = sentence  # 用新句子开始新的chunks

    if current_chunk:
        chunks.append(current_chunk)

    return chunks


# 更新进度条
class ProgressSignal(QObject):
    progress_updated = pyqtSignal(int)

class DesensitizeThread(QThread):
    finished = pyqtSignal(str)
    progress_updated = pyqtSignal(int)

    def __init__(self, input_file_path, output_file_path):
        super().__init__()
        self.input_file_path = input_file_path
        self.output_file_path = output_file_path
        self._is_running = True

    def run(self):
        try:
            file_paths = self._get_file_paths(self.input_file_path)
            total_files = len(file_paths)
            if total_files == 0:
                raise FileNotFoundError("没有检测到.docx文件！")

            for file_index, file_path in enumerate(file_paths):
                if not self._is_running:
                    break

                paragraphs = read_word_document(file_path)
                desensitized_paragraphs = []

                for paragraph in paragraphs:
                    chunks = split_text(paragraph, max_length=126)
                    for chunk in chunks:
                        desensitized_chunk = self._process_chunk(chunk)
                        desensitized_paragraphs.append(desensitized_chunk)

                desensitized_text = '\n'.join(desensitized_paragraphs)
                self._save_desensitized_file(file_path, desensitized_text)
                self._update_progress(file_index, total_files)

            self.finished.emit("脱敏完成. 结果保存至: " + self.output_file_path)
        except Exception as e:
            self.finished.emit("Error: " + str(e))
            
    def _process_chunk(self, chunk):
        try:
            chunk = re.sub(r'(编号)\d+|\w*公司', lambda m: '**公司' if '公司' in m.group() else m.group(1) + '**', chunk)
            entities = hanlp_ner_model(chunk)
            desensitized_chunk = chunk
            for entity in reversed(entities):
                # print(entity)
                entity_text, label, start, end = entity
                if label in ['NS', 'NT', 'NR']:  # 
                    print(entity_text, label, start, end)
                    desensitized_chunk = desensitized_chunk[:start] + '**' + desensitized_chunk[end:]
                    unit = re.search(r'(局|公司|工程|省|市|县|区)$', entity_text)
                    if unit:
                        replacement = '**' + unit.group()
                        desensitized_chunk = desensitized_chunk[:start] + replacement + desensitized_chunk[end:]
                        # if "社会主义" in desensitized_chunk:
                        #     print(desensitized_chunk)
                    else:
                        desensitized_chunk = desensitized_chunk[:start] + '**' + desensitized_chunk[end:]
            return desensitized_chunk
        except Exception as e:
            return chunk
    
    # def _process_chunk(self, chunk):
    #     try:
            # 识别并标记签发单位
            # ignored_ranges = [(m.start(), m.end()) for m in re.finditer(r'国务院|住房城乡建设部|财政部', chunk)]
    #         # 添加规则以忽略书名号内的内容和紧跟着中文括号的书名号
    #         ignored_ranges = [(m.start(), m.end()) for m in re.finditer(r'《.*?》(\（.*?\）)?', chunk)]

    #         # 用于检查一个位置是否在忽略的范围内
    #         def is_ignored_position(start, end):
    #             for ignored_start, ignored_end in ignored_ranges:
    #                 if start >= ignored_start and end <= ignored_end:
    #                     return True
    #             return False

    #         # 原有的脱敏逻辑
    #         chunk = re.sub(r'(编号)\d+|\w*公司', lambda m: '**公司' if '公司' in m.group() else m.group(1) + '**', chunk)
    #         entities = hanlp_ner_model(chunk)
    #         desensitized_chunk = chunk

    #         for entity in reversed(entities):
    #             entity_text, label, start, end = entity

    #             # 如果实体在忽略范围内，则不进行脱敏处理
    #             if is_ignored_position(start, end):
    #                 continue

    #             if label in ['NS', 'NT']:
    #                 unit = re.search(r'(局|公司|工程|省|市|县|区|社区)$', entity_text)
    #                 if unit:
    #                     replacement = '**' + unit.group()
    #                     desensitized_chunk = desensitized_chunk[:start] + replacement + desensitized_chunk[end:]
    #                 else:
    #                     desensitized_chunk = desensitized_chunk[:start] + '**' + desensitized_chunk[end:]
            
    #         return desensitized_chunk
    #     except Exception as e:
    #         print(e, chunk)
    #         return chunk


    def stop(self):
        self._is_running = False

    def _get_file_paths(self, input_path):
        if os.path.isfile(input_path):
            return [input_path]
        else:
            return [os.path.join(input_path, f) for f in os.listdir(input_path) if f.endswith('.docx') and not f.startswith('~$')]


    def _save_desensitized_file(self, original_file_path, desensitized_text):
        new_file_name = ''.join(random.choices(string.ascii_letters + string.digits, k=8)) + '.txt'
        output_file = os.path.join(self.output_file_path, new_file_name)
        write_to_text_file(output_file, desensitized_text)
        
        self._secondary_desensitization(output_file)

    def _update_progress(self, current_index, total_files):
        progress = int((current_index + 1) / total_files * 100)
        self.progress_updated.emit(progress)
        
    def _secondary_desensitization(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()

        # 将原始的脱敏逻辑应用于整个文档内容
        paragraphs = content.split('\n')
        desensitized_paragraphs = []

        for paragraph in paragraphs:
            chunks = split_text(paragraph, max_length=126)
            for chunk in chunks:
                desensitized_chunk = self._process_chunk(chunk)
                desensitized_paragraphs.append(desensitized_chunk)

        desensitized_text = '\n'.join(desensitized_paragraphs)

        # 重新写入脱敏后的内容
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(desensitized_text)


# 读取Word
def read_word_document(file_path):
    doc = Document(file_path)
    return [para.text for para in doc.paragraphs]

# 写入
def write_to_text_file(file_path, content):
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(content)


class MyApp(QWidget):
    
    trigger_select_input_file = pyqtSignal(str)
    trigger_select_output_file = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        
        self.trigger_select_input_file.connect(self.update_input_path)
        self.trigger_select_output_file.connect(self.update_output_path)

        # 创建垂直布局
        layout = QVBoxLayout()
        
        label1 = QLabel('请在转换之前将doc格式的文件转换为docx')
        layout.addWidget(label1)

        self.input1 = QLineEdit()
        self.input1.setPlaceholderText('请选择要脱敏的文件夹')
        layout.addWidget(self.input1)

        btn_upload = QPushButton('上传包含docx文件的文件夹')
        layout.addWidget(btn_upload)

        self.input2 = QLineEdit()
        self.input2.setPlaceholderText('请选择脱敏文件的输出位置')
        layout.addWidget(self.input2)

        btn_select_output = QPushButton('选择输出文件夹')
        layout.addWidget(btn_select_output)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        btn_start = QPushButton('开始脱敏')
        btn_stop = QPushButton('停止')
        layout.addWidget(btn_start)
        layout.addWidget(btn_stop)

        self.log_text = QTextEdit()
        layout.addWidget(self.log_text)
    
        self.setLayout(layout)

        btn_upload.clicked.connect(self.upload)
        btn_select_output.clicked.connect(self.select_output)
        btn_start.clicked.connect(self.start_process)
        btn_stop.clicked.connect(self.stop_process)

        # 文件路径
        self.input_file_path = ''
        self.output_file_path = ''
        
    def upload(self):
        file_path = QFileDialog.getExistingDirectory(self, '选择文件夹', '')

        if file_path:
            self.log_text.append(f"已选择文件夹：{file_path}")
            docx_files = [os.path.join(file_path, f) for f in os.listdir(file_path) if f.endswith('.docx') and not f.startswith('~$')]
            if docx_files:
                self.input_file_path = file_path  # 保存选择的文件夹路径
                self.input1.setText(file_path)
                self.log_text.append("文件夹内的docx文件有：\n" + "\n".join(docx_files))
            else:
                self.log_text.append("选择的文件夹内没有docx文件。")
        else:
            self.log_text.append("未选择文件夹。")


    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, '选择文件', '', 'Word 文档 (*.docx)')
        if file_path:
            self.trigger_select_input_file.emit(file_path)

    def select_output_file(self):
        file_path = QFileDialog.getSaveFileName(self, '选择输出文件', '', '文本文件 (*.txt)')[0]
        if file_path:
            self.trigger_select_output_file.emit(file_path)

    def select_output(self):
        file_path = QFileDialog.getExistingDirectory(self, '选择输出文件夹')
        if file_path:
            self.output_file_path = file_path
            self.input2.setText(file_path)
            self.log_text.append(f"输出文件夹选择完毕！为 {file_path}")
        else:
            self.log_text.append("未选择输出文件夹。")

    
    def start_process(self):
        if self.input_file_path and self.output_file_path:
            # 创建线程
            self.desensitize_thread = DesensitizeThread(self.input_file_path, self.output_file_path)

            self.desensitize_thread.progress_updated.connect(self.update_progress)
            self.desensitize_thread.finished.connect(self.process_finished)

            self.desensitize_thread.start()
            self.log_text.append("开始脱敏处理。")
            
    def update_input_path(self, path):
        self.input1.setText(path)

    def update_output_path(self, path):
        self.input2.setText(path)


    def stop_process(self):
        if self.desensitize_thread and self.desensitize_thread.isRunning():
            self.desensitize_thread.stop()  
            self.progress.setValue(0)
            self.log_text.append("任务已停止。")
            
    def update_progress(self, progress):
        self.progress.setValue(progress)

    def process_finished(self, message):
        self.log_text.setText(message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    ex.show()
    sys.exit(app.exec_())
    
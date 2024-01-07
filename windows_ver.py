import sys
import re
import os
import random
import string
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QProgressBar, QTextEdit, QFileDialog
from PyQt5.QtCore import QThread, pyqtSignal, QObject
from pathlib import Path
from PyPDF2 import PdfReader
import pythoncom
from win32com import client
from docx import Document
import hanlp

if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

huggingface_cache_path = os.path.join(application_path, 'huggingface')
os.environ['TRANSFORMERS_CACHE'] = huggingface_cache_path

model_path = os.path.join(application_path, 'ner_bert_base_msra_20211227_114712')
hanlp_ner_model = hanlp.load(model_path)

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
                current_chunk = sentence

    if current_chunk:
        chunks.append(current_chunk)

    return chunks

def read_pdf_document(file_path):
    with open(file_path, 'rb') as file:
        reader = PdfReader(file)
        num_pages = len(reader.pages)
        paras = [reader.pages[page].extract_text() + '\n' for page in range(num_pages)]
        if all(para.isspace() for para in paras):
            return 'PDF文件为空'
        else:
            return paras
        
def deal_path(path):
    path = path.replace('\\', '/')
    return path 

def read_doc_document(file_path):
    word = client.Dispatch("Word.Application")
    word.Visible = False  # Word应用程序在后台运行，不显示界面
    file_path = file_path.replace('/', '\\')
    doc = word.Documents.Open(file_path)
    paragraphs = []
    temp_para = ''
    for para in doc.Paragraphs:
        if para.Range.Text.strip().isspace():
            continue
        temp_para += para.Range.Text.strip() + '\n'  # 使用'\n'来分隔每个段落
        if len(temp_para) > 126:
            paragraphs.append(temp_para)
            temp_para = ''
    if temp_para:  # 如果最后还有剩余的段落，也添加到列表中
        paragraphs.append(temp_para)
    doc.Close(False)
    word.Quit()
    return paragraphs


#检查文件格式并且返回段落文本
def read_file_context(file_path):
    if file_path.endswith('.docx'):
        paragraphs = read_word_document(file_path)
    elif file_path.endswith('.pdf'):
        paragraphs = read_pdf_document(file_path)
    elif file_path.endswith('.doc'):
        print('正在处理doc文件')
        paragraphs = read_doc_document(file_path)
    return paragraphs

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
                raise FileNotFoundError("没有检测到可以于脱敏的文件！")

            for file_index, file_path in enumerate(file_paths):
                file_path = deal_path(file_path)
                if not self._is_running:
                    self._clean_up_and_exit()  # 清理并退出线程
                    return
                try:
                    paragraphs = read_file_context(file_path)
                    if paragraphs == 'PDF文件为空':
                        continue
                except Exception as e:
                    continue
                desensitized_paragraphs = []
                print(f'正在脱敏文件：{file_path}')
                for paragraph in paragraphs:
                    if not self._is_running:
                        self._clean_up_and_exit()  # 清理并退出线程
                        return
                    chunks = split_text(paragraph, max_length=126)
                    for chunk in chunks:
                        desensitized_chunk = self._process_chunk(chunk)
                        #执行2次脱敏
                        desensitized_chunk = self._process_chunk(desensitized_chunk)
                        desensitized_paragraphs.append(desensitized_chunk)

                desensitized_text = ''.join(desensitized_paragraphs)
                self._save_desensitized_file(file_path, desensitized_text)
                self._update_progress(file_index, total_files)
                
                if not self._is_running:
                    self._clean_up_and_exit()  # 清理并退出线程
                    return

            self.finished.emit("脱敏完成. 结果保存至: " + self.output_file_path)
        except Exception as e:
            self.finished.emit("Error: " + str(e))

    def _process_chunk(self, chunk):
        try:
            chunk = re.sub(r'(编号)\d+|\w*公司', lambda m: '*公司' if '公司' in m.group() else m.group(1) + '*', chunk)
            entities = hanlp_ner_model(chunk)
            desensitized_chunk = chunk
            for entity in reversed(entities):
                entity_text, label, start, end = entity
                if label in ['NS', 'NT', 'NR']:
                    unit = re.search(r'(局|公司|工程|省|市|县|区|社区)$', entity_text)
                    if unit:
                        replacement = '*' + unit.group()
                        desensitized_chunk = desensitized_chunk[:start] + replacement + desensitized_chunk[end:]
                    else:
                        desensitized_chunk = desensitized_chunk[:start] + '*' + desensitized_chunk[end:]
            return desensitized_chunk
        except Exception as e:
            return chunk

    def stop(self):
        self._is_running = False

    def _clean_up_and_exit(self):
        if hasattr(self, 'file'):
            self.file.close()
            del self.file
        self.finished.emit("任务已停止。")

    def _get_file_paths(self, input_path):
        file_paths = []
        for root, dirs, files in os.walk(input_path):
            for file in files:
                if file.lower().endswith(('doc', 'docx', 'pdf')) and not file.startswith('~$') and not file.startswith('.'):
                    file_path = os.path.join(root, file)
                    file_paths.append(file_path)
        return file_paths



    def _save_desensitized_file(self, original_file_path, desensitized_text):
        new_file_name = ''.join(random.choices(string.ascii_letters + string.digits, k=8)) + '_' + os.path.basename(original_file_path) + '.txt'
        output_file = os.path.join(self.output_file_path, new_file_name)
        write_to_text_file(output_file, desensitized_text)

    def _update_progress(self, current_index, total_files):
        progress = int((current_index + 1) / total_files * 100)
        self.progress_updated.emit(progress)

def read_word_document(file_path):
    doc = Document(file_path)
    return [para.text + '\n' for para in doc.paragraphs]

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

        layout = QVBoxLayout()
        
        label1 = QLabel('目前可以脱敏的文件类型有：docx, doc, pdf')
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

        self.input_file_path = ''
        self.output_file_path = ''
        
    def upload(self):
        file_path = QFileDialog.getExistingDirectory(self, '选择文件夹', '')

        if file_path:
            self.log_text.append(f"已选择文件夹：{file_path}")
            docx_files = [os.path.join(root, f) 
                        for root, dirs, files in os.walk(file_path) 
                        for f in files if f.lower().endswith(('.docx', '.pdf', '.doc')) and not f.startswith('~$') and not f.startswith('.')]
            if docx_files:
                self.input_file_path = file_path
                self.input1.setText(file_path)
                self.log_text.append("文件夹内可以脱敏的文件有：" + str(len(docx_files)) + "个 \n" + "\n".join(docx_files))
            else:
                self.log_text.append("选择的文件夹内没有docx文件。")
        else:
            self.log_text.append("未选择文件夹。")


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


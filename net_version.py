import sys
import re
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QProgressBar, QTextEdit, QFileDialog
from PyQt5.QtCore import QThread, pyqtSignal, QObject
import os
import requests
import random
import string
from docx import Document


# 新的脱敏函数，使用API进行实体识别和脱敏
def desensitize_with_api(text):
    try:
        url = API_URL
        headers = {
            "token": API_TOKEN  # 直接使用 token 而不是 Bearer 方案
        }
        data = {
            "text": text
        }

        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()

        result = response.json()
        return result
    except Exception as e:
        return {"error": str(e)}



# 创建一个信号类以用于更新进度条
class ProgressSignal(QObject):
    progress_updated = pyqtSignal(int)

# 创建一个线程类以执行脱敏任务
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

                desensitized_text = self._process_file(file_path)
                self._save_desensitized_file(file_path, desensitized_text)
                self._update_progress(file_index, total_files)

            self.finished.emit("脱敏完成. 结果保存至: " + self.output_file_path)
        except Exception as e:
            self.finished.emit("Error: " + str(e))

    def stop(self):
        self._is_running = False

    def _get_file_paths(self, input_path):
        if os.path.isfile(input_path):
            return [input_path]
        else:
            return [os.path.join(input_path, f) for f in os.listdir(input_path) if f.endswith('.docx') and not f.startswith('~$')]

    def _process_file(self, file_path):
        paragraphs = read_word_document(file_path)
        desensitized_paragraphs = [
            desensitize_with_api_and_regex(paragraph, desensitize_with_api(paragraph))
            for paragraph in paragraphs
        ]
        return '\n'.join(desensitized_paragraphs)

    def _save_desensitized_file(self, original_file_path, desensitized_text):
        base_name = os.path.splitext(os.path.basename(original_file_path))[0]
        new_file_name = base_name + '_' + ''.join(random.choices(string.ascii_letters + string.digits, k=8)) + '.txt'
        output_file = os.path.join(self.output_file_path, new_file_name)
        write_to_text_file(output_file, desensitized_text)

    def _update_progress(self, current_index, total_files):
        progress = int((current_index + 1) / total_files * 100)
        self.progress_updated.emit(progress)



def desensitize_with_api_and_regex(text, api_response):
    desensitized_text = text
    print(api_response)
    # API 返回的实体数据
    entities = api_response.get('data', {}).get('ner/msra', [])
    
    for ent in reversed(entities):
        if len(ent) == 4:
            entity_text, label, start, end = ent
            # 只处理特定类型的实体
            if label in ['PERSON', 'LOCATION']:  # 根据您的需要调整实体类型
                # 用正则表达式找到单位（如“局”，“公司”等）并保留
                unit = re.search(r'(局|公司|工程|省|市|县|区|街道|社区|小区|花园|苑)$', entity_text)
                if unit:
                    replacement = '**' + unit.group()
                    desensitized_text = desensitized_text[:start] + replacement + desensitized_text[end:]
                else:
                    # 如果没有单位，就替换整个实体
                    desensitized_text = desensitized_text[:start] + '**' + desensitized_text[end:]

    return desensitized_text


# 读取Word文档的函数
def read_word_document(file_path):
    doc = Document(file_path)
    return [para.text for para in doc.paragraphs]  # 直接返回段落文本的列表

# 写入文本文件的函数
def write_to_text_file(file_path, content):
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(content)  # 确保 content 是正确格式化的字符串


class MyApp(QWidget):
    
    trigger_select_input_file = pyqtSignal(str)
    trigger_select_output_file = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        
        self.trigger_select_input_file.connect(self.update_input_path)
        self.trigger_select_output_file.connect(self.update_output_path)
        
        # 初始化hanlp模型
        # self.nlp = hanlp.load(hanlp.pretrained.tok.LARGE_ALBERT_BASE)
        # self.nlp = hanlp.load(hanlp.pretrained.ner.MSRA_NER_BERT_BASE_ZH)
        global API_URL, API_TOKEN
        API_URL = "http://comdo.hanlp.com/hanlp/v21/ner/ner"
        API_TOKEN = "b48ae41e0f5941b0a495ec2e1543aa931701144867858token"

        # 创建垂直布局
        layout = QVBoxLayout()
        

        # 第一栏：提示信息
        label1 = QLabel('请在转换之前将doc格式的文件转换为docx')
        layout.addWidget(label1)

        # 第二栏：输入框和上传按钮
        self.input1 = QLineEdit()
        self.input1.setPlaceholderText('请选择要脱敏的文件或者文件夹')
        layout.addWidget(self.input1)

        btn_upload = QPushButton('上传文件或文件夹')
        layout.addWidget(btn_upload)

        # 第三栏：输出位置选择
        self.input2 = QLineEdit()
        self.input2.setPlaceholderText('请选择脱敏文件的输出位置')
        layout.addWidget(self.input2)

        btn_select_output = QPushButton('选择输出文件夹')
        layout.addWidget(btn_select_output)

        # 进度条
        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        # 开始和停止按钮
        btn_start = QPushButton('开始脱敏')
        btn_stop = QPushButton('停止')
        layout.addWidget(btn_start)
        layout.addWidget(btn_stop)

        # 日志显示框
        self.log_text = QTextEdit()
        layout.addWidget(self.log_text)
        # 设置按钮的样式
    

        # 设置布局
        self.setLayout(layout)

        # 按钮的点击事件
        btn_upload.clicked.connect(self.upload)
        btn_select_output.clicked.connect(self.select_output)
        btn_start.clicked.connect(self.start_process)
        btn_stop.clicked.connect(self.stop_process)

        # 文件路径
        self.input_file_path = ''
        self.output_file_path = ''
        
    def upload(self):
        # 使用 QFileDialog.getExistingDirectory 而不是 getOpenFileName
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

            # 连接信号和槽函数
            self.desensitize_thread.progress_updated.connect(self.update_progress)
            self.desensitize_thread.finished.connect(self.process_finished)

            # 启动线程
            self.desensitize_thread.start()
            self.log_text.append("开始脱敏处理。")
            
    def update_input_path(self, path):
        self.input1.setText(path)

    def update_output_path(self, path):
        self.input2.setText(path)


    def stop_process(self):
        if self.desensitize_thread and self.desensitize_thread.isRunning():
            self.desensitize_thread.stop()  # 使用新的stop方法
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

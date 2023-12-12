import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QFileDialog, QMessageBox
from cryptography.fernet import Fernet

class DecryptApp(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('文件解密工具')
        layout = QVBoxLayout()

        self.key_path_input = QLineEdit(self)
        self.key_path_input.setPlaceholderText('上传密钥文件')
        layout.addWidget(self.key_path_input)

        key_upload_button = QPushButton('上传密钥文件', self)
        key_upload_button.clicked.connect(self.upload_key_file)
        layout.addWidget(key_upload_button)

        self.folder_path_input = QLineEdit(self)
        self.folder_path_input.setPlaceholderText('选择需要解密的文件夹')
        layout.addWidget(self.folder_path_input)

        folder_upload_button = QPushButton('选择文件夹', self)
        folder_upload_button.clicked.connect(self.upload_folder)
        layout.addWidget(folder_upload_button)

        decrypt_button = QPushButton('开始解密', self)
        decrypt_button.clicked.connect(self.decrypt_folder)
        layout.addWidget(decrypt_button)

        self.setLayout(layout)

    def upload_key_file(self):
        key_file_path, _ = QFileDialog.getOpenFileName(self, '选择密钥文件', '', 'Key File (*.key)')
        if key_file_path:
            self.key_path_input.setText(key_file_path)

    def upload_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, '选择文件夹', '')
        if folder_path:
            self.folder_path_input.setText(folder_path)

    def decrypt_folder(self):
        folder_path = self.folder_path_input.text()
        key_file_path = self.key_path_input.text()

        if folder_path and key_file_path:
            try:
                with open(key_file_path, 'rb') as key_file:
                    key = key_file.read()
                fernet = Fernet(key)

                for file_name in os.listdir(folder_path):
                    if file_name.endswith('.enc'):
                        file_path = os.path.join(folder_path, file_name)
                        with open(file_path, 'rb') as file:
                            encrypted_data = file.read()

                        decrypted_data = fernet.decrypt(encrypted_data)

                        decrypted_file_path = os.path.splitext(file_path)[0] + '_decrypted.txt'
                        with open(decrypted_file_path, 'wb') as file:
                            file.write(decrypted_data)

                QMessageBox.information(self, '解密成功', '文件夹中的所有文件已解密')
            except Exception as e:
                QMessageBox.warning(self, '解密失败', '解密过程中发生错误。\n' + str(e))
        else:
            QMessageBox.warning(self, '错误', '请上传密钥文件和选择需要解密的文件夹')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DecryptApp()
    ex.show()
    sys.exit(app.exec_())

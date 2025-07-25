import winreg
import os
import json
import ctypes
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                           QListWidget, QMessageBox, QPushButton, QHBoxLayout,
                           QStyle, QListWidgetItem, QLabel)
from PyQt5.QtCore import Qt, QPropertyAnimation, QPoint
from PyQt5.QtGui import QTransform, QIcon, QFont, QColor

def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    try:
        if not is_admin():
            # 重新以管理员身份运行
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            sys.exit()
        return True
    except Exception as e:
        print(f"无法获取管理员权限: {e}")
        return False

class ContextMenuList(QListWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet("""
            QListWidget {
                border: 1px solid #cccccc;
                border-radius: 5px;
                background-color: #ffffff;
                padding: 5px;
            }
            QListWidget::item {
                height: 30px;
                border-bottom: 1px solid #eeeeee;
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #e6f3ff;
                color: #000000;
            }
            QListWidget::item:hover {
                background-color: #f5f5f5;
            }
        """)

class StyledButton(QPushButton):
    def __init__(self, text, icon_name=None):
        super().__init__(text)
        if icon_name:
            self.setIcon(self.style().standardIcon(icon_name))
        self.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 6px 12px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #e6f3ff;
                border-color: #1890ff;
            }
            QPushButton:pressed {
                background-color: #1890ff;
                color: white;
            }
        """)

# 注册表路径
BACKGROUND_SHELL = r'Directory\Background\shell'

# 系统内置应用列表
SYSTEM_APPS = {
    'cmd',                  # 命令提示符
    'Powershell',          # PowerShell
    'runas',               # 以管理员身份运行
    'WindowsTerminal',     # Windows 终端
    'SystemSettings',      # 系统设置
    'Microsoft.Windows.',  # Windows 商店应用前缀
}

def is_system_app(app_name):
    """检查是否为系统内置应用"""
    return any(app_name.startswith(sys_app) for sys_app in SYSTEM_APPS)

# 备份文件
def backup_menu_items():
    backup = {}
    try:
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, BACKGROUND_SHELL)
        i = 0
        while True:
            try:
                subkey_name = winreg.EnumKey(key, i)
                if not is_system_app(subkey_name):  # 只备份非系统应用
                    subkey_path = f'{BACKGROUND_SHELL}\\{subkey_name}'
                    subkey = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, subkey_path)
                    backup[subkey_name] = {}
                    j = 0
                    while True:
                        try:
                            name, value, _ = winreg.EnumValue(subkey, j)
                            backup[subkey_name][name] = value
                            j += 1
                        except OSError:
                            break
                    subkey.Close()
                i += 1
            except OSError:
                break
        key.Close()

        # 保存备份到文件
        backup_file = get_resource_path('menu_backup.json')
        with open(backup_file, 'w', encoding='utf-8') as f:
            json.dump(backup, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"备份失败: {str(e)}")
        return False

def list_menu_items():
    items = []
    try:
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, BACKGROUND_SHELL)
        i = 0
        while True:
            try:
                subkey_name = winreg.EnumKey(key, i)
                if not is_system_app(subkey_name):  # 只添加非系统应用
                    items.append(subkey_name)
                i += 1
            except OSError:
                break
        key.Close()
    except Exception as e:
        print(f'获取菜单失败: {e}')
        return []
    return items

def delete_menu_item(item_name):
    if is_system_app(item_name):
        print(f'跳过系统应用 {item_name}')
        return False

    path = f'{BACKGROUND_SHELL}\\{item_name}'
    def delete_key_recursive(key_path):
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_ALL_ACCESS)
            info = winreg.QueryInfoKey(key)
            for i in range(info[0]):
                subkey = winreg.EnumKey(key, 0)
                delete_key_recursive(key_path + '\\' + subkey)
            winreg.CloseKey(key)
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
        except Exception as e:
            print(f'删除子键失败: {e}')
            return False
        return True
    if delete_key_recursive(path):
        print(f'删除 {item_name} 成功。')
    else:
        print(f'删除 {item_name} 失败，请确保以管理员权限运行脚本。')

def restore_menu_items(backup_file='menu_backup.json'):
    if not is_admin():
        QMessageBox.critical(None, '错误', '需要管理员权限才能恢复菜单项！\n请以管理员身份运行。')
        return False

    backup_file = get_resource_path(backup_file)
    if not os.path.exists(backup_file):
        QMessageBox.critical(None, '错误', '备份文件不存在。')
        return False

    try:
        with open(backup_file, 'r', encoding='utf-8') as f:
            backup = json.load(f)
    except Exception as e:
        QMessageBox.critical(None, '错误', f'读取备份文件失败: {str(e)}')
        return False

    success = True
    failed_items = []

    for item, values in backup.items():
        if is_system_app(item):  # 跳过系统应用的恢复
            continue

        path = f'{BACKGROUND_SHELL}\\{item}'
        try:
            # 如果键已存在，先删除
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, path)
            except WindowsError:
                pass

            key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, path)
            for name, value in values.items():
                try:
                    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
                except Exception as e:
                    print(f'设置值失败 {item}.{name}: {str(e)}')
                    failed_items.append(f'{item}.{name}')
            winreg.CloseKey(key)
        except Exception as e:
            print(f'恢复 {item} 失败: {str(e)}')
            failed_items.append(item)
            success = False

    if failed_items:
        error_msg = '以下项目恢复失败：\n' + '\n'.join(failed_items)
        QMessageBox.warning(None, '部分恢复失败', error_msg)
    elif success:
        QMessageBox.information(None, '成功', '所有菜单项已成功恢复。')

    return success

class ContextMenuManager(QMainWindow):
    def __init__(self):
        super().__init__()
        if not is_admin():
            QMessageBox.warning(None, '权限提示',
                              '程序未以管理员权限运行！\n某些操作可能无法执行，请右键选择"以管理员身份运行"。')

        self.setWindowTitle('右键菜单管理器')
        self.setGeometry(300, 300, 800, 500)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5;
            }
        """)

        # 创建中心部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # 添加标题
        title_label = QLabel('Windows 右键菜单管理')
        title_label.setFont(QFont('Arial', 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # 创建列表控件
        self.list_widget = ContextMenuList()
        layout.addWidget(self.list_widget)

        # 创建按钮布局
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        layout.addLayout(button_layout)

        # 创建按钮
        self.refresh_button = StyledButton('刷新列表', QStyle.SP_BrowserReload)
        self.refresh_button.clicked.connect(self.refresh_list)
        button_layout.addWidget(self.refresh_button)

        self.backup_button = StyledButton('备份', QStyle.SP_DialogSaveButton)
        self.backup_button.clicked.connect(self.backup_current_items)
        button_layout.addWidget(self.backup_button)

        self.delete_button = StyledButton('删除选中', QStyle.SP_DialogDiscardButton)
        self.delete_button.clicked.connect(self.delete_selected)
        button_layout.addWidget(self.delete_button)

        self.restore_button = StyledButton('恢复', QStyle.SP_BrowserReload)
        self.restore_button.clicked.connect(self.restore)
        button_layout.addWidget(self.restore_button)

        # 添加状态栏
        self.statusBar().showMessage('准备就绪')

        # 加载菜单项
        self.load_menu_items()

    def load_menu_items(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, BACKGROUND_SHELL)
            i = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    if not is_system_app(subkey_name):  # 只添加非系统应用
                        self.list_widget.addItem(subkey_name)
                    i += 1
                except OSError:
                    break
            key.Close()
        except Exception as e:
            QMessageBox.critical(self, '错误', f'加载菜单项失败: {str(e)}')

    def backup_current_items(self):
        if backup_menu_items():
            QMessageBox.information(self, '成功', '菜单项已成功备份')
        else:
            QMessageBox.critical(self, '错误', '备份失败')

    def refresh_list(self):
        self.list_widget.clear()
        items = list_menu_items()
        for item in items:
            self.list_widget.addItem(item)

    def delete_selected(self):
        if not is_admin():
            QMessageBox.critical(self, '错误', '需要管理员权限才能删除菜单项！\n请右键选择"以管理员身份运行"。')
            return

        try:
            selected_item = self.list_widget.currentItem()
            if not selected_item:
                QMessageBox.warning(self, '警告', '请选择要删除的菜单项。')
                return

            item_name = selected_item.text()
            reply = QMessageBox.question(self, '确认删除',
                                       f'确定要删除"{item_name}"吗？\n此操作不可恢复！',
                                       QMessageBox.Yes | QMessageBox.No)

            if reply == QMessageBox.Yes:
                delete_menu_item(item_name)
                self.refresh_list()
                self.statusBar().showMessage(f'已删除菜单项：{item_name}')
        except Exception as e:
            QMessageBox.critical(self, '错误', f'删除失败：{str(e)}')

    def restore(self):
        if restore_menu_items():
            self.refresh_list()
            self.statusBar().showMessage('菜单项已恢复')
        else:
            self.statusBar().showMessage('恢复失败')

def main():
    # 确保以管理员权限运行
    run_as_admin()

    app = QApplication([])
    window = ContextMenuManager()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()

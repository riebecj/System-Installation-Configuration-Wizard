import os
import subprocess
from PyQt5 import QtCore, QtGui, QtWidgets
import getpass
import ctypes
from ctypes import wintypes
import sys
import socket
import shutil

_GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
_GetShortPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
_GetShortPathNameW.restype = wintypes.DWORD
try:
    temp_path = sys._MEIPASS
except:
    temp_path = None

try:
    img_path = os.path.join(temp_path, 'items')
except:
    img_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'items')


def get_short_path_name(long_name):
    output_buf_size = 0
    while True:
        output_buf = ctypes.create_unicode_buffer(output_buf_size)
        needed = _GetShortPathNameW(long_name, output_buf, output_buf_size)
        if output_buf_size >= needed:
            return output_buf.value
        else:
            output_buf_size = needed


def install_packages():
    working_dir = os.getcwd()
    subprocess.call(['python', '{}'.format(os.path.join(working_dir, 'upgrade_pip.py'))])
    import pip



    installed = []
    for i in pip.get_installed_distributions():
        installed.append(str(i))

    dirpath = os.getcwd()
    for module in os.listdir(os.path.join(dirpath, 'packages')):
        package = module.split('-')[0]
        packPath = os.path.join(dirpath, 'packages')

        if any(package in s for s in installed):
            pip.main(['install', '--upgrade', os.path.join(packPath, module)])

        else:
            pip.main(['install', os.path.join(packPath, module)])


class QIComboBox(QtWidgets.QComboBox):
    def __init__(self, parent=None):
        super(QIComboBox, self).__init__(parent)


class Page1(QtWidgets.QWizardPage):
    def __init__(self, parent=None):
        super(Page1, self).__init__(parent)
        self.build()
        self.set_layout()

    def build(self):
        welcome_font = QtGui.QFont('TimesNewRoman', 50)
        intro_font = QtGui.QFont('TimesNewRoman', 12)
        self.Title = QtWidgets.QLabel('Welcome!', self)
        self.Title.setFont(welcome_font)

        self.intro = QtWidgets.QLabel('\nThis wizard will help install the AEA-GALE Data Integration Tool (ADIT) on '
                                      'your machine and/or configure your GALE client to connect to the PostgreSQL '
                                      'database on COLLAB.')
        self.intro.setFont(intro_font)
        self.intro.setWordWrap(True)

        self.img = QtWidgets.QLabel(self)
        pixmap = QtGui.QPixmap(os.path.join(img_path, 'logo.jpg'))
        self.img.setPixmap(pixmap)

        self.frame = QtWidgets.QFrame()
        self.frame.resize(200, 200)

    def set_layout(self):
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.Title)
        layout.addWidget(self.intro)
        layout.addWidget(self.frame)
        layout2 = QtWidgets.QHBoxLayout()
        layout2.addWidget(self.img)
        layout2.addLayout(layout)
        self.setLayout(layout2)


class Page2(QtWidgets.QWizardPage):
    def __init__(self, parent=None):
        super(Page2, self).__init__(parent)
        self.parent = parent
        self.gale_dir = None

    def initializePage(self):
        self.build()
        self.set_layout()
        self.find_gale_install_dir()
        self.entry.setText(self.gale_dir)
        self.server.setFocus()

    def build(self):
        self.img = QtWidgets.QLabel(self)
        pixmap = QtGui.QPixmap(os.path.join(img_path, 'logo.jpg'))
        self.img.setPixmap(pixmap)
        text_font = QtGui.QFont('TimesNewRoman', 10)
        self.label = QtWidgets.QLabel('Ensure the directory displayed below is the current GALE profile installation '
                                      'directory (the directory below should contain the folder "GALE_Application"). '
                                      'If it has changed, please Browse to select the current profile '
                                      'directory, otherwise click Next. \n\nIf you are unsure, please contact your '
                                      'Administrator.')
        self.label.setFont(text_font)
        self.label.setWordWrap(True)

        self.frame = QtWidgets.QFrame()
        self.frame.resize(200,200)

        self.entry = QtWidgets.QLineEdit()
        self.browse = QtWidgets.QPushButton('Browse')
        self.browse.clicked.connect(self.file_dialog)

        self.label2 = QtWidgets.QLabel('Enter the IP or Domain Name for your available GALE PostgreSQL Server.\n\nNOTE:'
                                       'If Domain Name different than RELAUSGALE or localhost, it will still say '
                                       'invalid, but the installer will still install correctly.')
        self.label2.setWordWrap(True)
        self.label2.setFont(text_font)
        self.server = QtWidgets.QLineEdit(self)
        self.server.textChanged.connect(self.validate)
        self.status = QtWidgets.QLabel()
        self.status.setText('Invalid')
        self.status.setObjectName('status')
        self.status.setStyleSheet('QLabel#status {color: red}')

    def set_layout(self):
        layout1 = QtWidgets.QHBoxLayout()
        layout1.addWidget(self.entry)
        layout1.addWidget(self.browse)

        layout3 = QtWidgets.QHBoxLayout()
        layout3.addWidget(self.server)
        layout3.addWidget(self.status)

        layout2 = QtWidgets.QVBoxLayout()
        layout2.addWidget(self.label)
        layout2.addLayout(layout1)
        layout2.addWidget(self.frame)
        layout2.addWidget(self.label2)
        layout2.addLayout(layout3)
        layout2.addWidget(self.frame)

        layout = QtWidgets.QHBoxLayout()
        layout.addWidget(self.img)
        layout.addLayout(layout2)
        self.setLayout(layout)

    def find_gale_install_dir(self):
        if os.path.isdir(r"C:\GALEData"):
            self.gale_dir = r'C:\GALEData'
        elif 'GALE_Application' in os.listdir(r'C:\Users\{}\AppData\Local'.format(getpass.getuser())):
            self.gale_dir = r'C:\Users\{}\AppData\Local'.format(getpass.getuser())
        else:
            QtWidgets.QMessageBox.critical(self, 'GALE Install Error', 'GALE Installation directory cannot be '
                                                   'found. Re-install GALE and Relaunch this installation.',
                                                   QtWidgets.QMessageBox.Close, QtWidgets.QMessageBox.Close)
            sys.exit(0)

    def file_dialog(self):
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select GALE Profile Install Directory',
                                                               self.gale_dir,
                                                               QtWidgets.QFileDialog.ShowDirsOnly)
        if os.path.isdir(directory):
            self.entry.setText(directory.replace('/', '\\'))
            self.gale_dir = directory.replace('/', '\\')

    def validatePage(self):
        if 'GALE_Application' in os.listdir(self.gale_dir) or \
                all(i in os.listdir(self.gale_dir) for i in ['buckets', 'config', 'FST', 'log', 'tools']):
            if self.server.text() != '':
                return True
            else:
                QtWidgets.QMessageBox.information(self, 'Server Error',
                                                  'Please specify a server IP or Domain Name.',
                                                  QtWidgets.QMessageBox.Close, QtWidgets.QMessageBox.Close)
                return False
        else:
            QtWidgets.QMessageBox.information(self, 'Directory Error', 'The specified directory does not contain a GALE'
                                            ' profile installation or the installation was not completed properly. '
                                            'Please select the appropriate directory or contact your network '
                                            'Administrator for guidance.',
                                              QtWidgets.QMessageBox.Close, QtWidgets.QMessageBox.Close)
            return False

    def validate(self):
        try:
            if self.server.text() == 'localhost':
                self.status.setStyleSheet('QLabel#status {color: green}')
                self.status.setText('Valid')
            elif self.server.text() == 'RELAUSGALE' or self.server.text() == 'relausgale':
                self.status.setStyleSheet('QLabel#status {color: green}')
                self.status.setText('Valid')
            else:
                socket.inet_aton(self.server.text())
                self.status.setStyleSheet('QLabel#status {color: green}')
                self.status.setText('Valid')
        except socket.error:
            self.status.setText('Invalid')
            self.status.setStyleSheet('QLabel#status {color: red}')


class Page3(QtWidgets.QWizardPage):
    selectionChanged = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        super(Page3, self).__init__(parent)
        self.parent = parent
        self.install_version = None

    def initializePage(self):
        try:
            self.build()
            self.set_layout()
            self.b1.setChecked(True)
        except:
            print(sys.exc_info()[1])

    def build(self):
        self.img = QtWidgets.QLabel(self)
        pixmap = QtGui.QPixmap(os.path.join(img_path, 'logo.jpg'))
        self.img.setPixmap(pixmap)

        text_font = QtGui.QFont('TimesNewRoman', 10)

        self.label1 = QtWidgets.QLabel('ADIT for EDES 1.0/2.0: ')
        self.label2 = QtWidgets.QLabel('EWDS Developed ADIT: ')
        self.label3 = QtWidgets.QLabel('Decode, Process, and Extract Post-Flight data from raw binary file from DMD. '
                                       'Will output compressed CSV, DSR, and Upload data to local/networked PostgreSQL'
                                       ' Server.')
        self.label3.setWordWrap(True)
        self.label4 = QtWidgets.QLabel('Extract Post-Flight data from previously decoded mission files from either '
                                       'JMPS or ADIT for EDES tool. Will output compressed CSV, DSR, and Upload data '
                                       'to local/networked PostgreSQL Server.')
        self.label4.setWordWrap(True)
        self.label1.setFont(text_font)
        self.label2.setFont(text_font)
        self.label3.setFont(text_font)
        self.label4.setFont(text_font)
        self.box = QtWidgets.QGroupBox()
        self.box.setTitle('Select ADIT Version to Install')
        self.box.setLayout(QtWidgets.QVBoxLayout())

        self.group_adit = QtWidgets.QButtonGroup()
        self.b1 = QtWidgets.QRadioButton('ADIT for EDES 1.0/2.0')
        self.b2 = QtWidgets.QRadioButton('EWDS Developed ADIT')
        self.group_adit.addButton(self.b1)
        self.group_adit.addButton(self.b2)
        self.b1.toggled.connect(self.selection)

        self.box.layout().addWidget(self.b1)
        self.box.layout().addWidget(self.b2)

    def selection(self):
        checked_button = self.group_adit.checkedButton()
        self.install_version = checked_button.text()

    def set_layout(self):
        self.nl1 = QtWidgets.QHBoxLayout()
        self.nl2 = QtWidgets.QHBoxLayout()
        self.nl1.addWidget(self.label1)
        self.nl1.addWidget(self.label3)
        self.nl2.addWidget(self.label2)
        self.nl2.addWidget(self.label4)

        self.layout = QtWidgets.QVBoxLayout()
        self.layout.addWidget(self.box)
        self.layout.addLayout(self.nl1)
        self.layout.addLayout(self.nl2)

        self.layout_main = QtWidgets.QHBoxLayout()
        self.layout_main.addWidget(self.img)
        self.layout_main.addLayout(self.layout)
        self.setLayout(self.layout_main)


class Page4(QtWidgets.QWizardPage):
    def __init__(self, parent=None):
        super(Page4, self).__init__(parent)
        self.parent = parent
        self.complete = False

    def initializePage(self):
        self.build()
        self.set_layout()
        self.thread.pbar_signal.connect(self.bar.setValue)
        self.thread.pbar_action.connect(self.label.setText)
        self.cancel.clicked.connect(self.stop_thread)
        self.thread.done.connect(self.update_gui)
        self.thread.log_signal.connect(self.update_log)
        self.thread.message.connect(self.message)
        if self.parent.Page3.install_version == "ADIT for EDES 1.0/2.0":
            self.label.setText('Verifying ADIT Installation...')
        else:
            self.label.setText('Upgrading Pip...')

    def update_log(self, string):
        if string != '':
            self.log.append(string)

    def message(self, string):
        title, message = string.split('-')
        QtWidgets.QMessageBox.warning(self.parent, title, message, QtWidgets.QMessageBox.Close,
                                      QtWidgets.QMessageBox.Close)
        if title == 'Permission Error':
            sys.exit(14)

    def build(self):
        self.img = QtWidgets.QLabel(self)
        pixmap = QtGui.QPixmap(os.path.join(img_path, 'logo.jpg'))
        self.img.setPixmap(pixmap)

        self.label = QtWidgets.QLabel()
        self.bar = QtWidgets.QProgressBar(self)
        self.cancel = QtWidgets.QPushButton('Cancel')

        self.log = QtWidgets.QTextEdit()

        self.frame = QtWidgets.QFrame()
        self.frame.resize(200, 200)
        self.thread = WorkThread(self)

    def update_gui(self):
        self.completeChanged.emit()

    def stop_thread(self):
        if self.thread.isRunning():
            self.thread.stop()
        self.completeChanged.emit()

    def isComplete(self, x=None):
        if self.bar.value() == 100:
            self.complete = True

        return self.complete

    def set_layout(self):
        layout1 = QtWidgets.QHBoxLayout()
        layout1.addWidget(self.bar)
        layout1.addWidget(self.cancel)

        layout2 = QtWidgets.QVBoxLayout()
        layout2.addWidget(self.frame)
        layout2.addWidget(self.frame)
        layout2.addWidget(self.label)
        layout2.addLayout(layout1)
        layout2.addWidget(self.log)
        layout2.addWidget(self.frame)

        layout_main = QtWidgets.QHBoxLayout()
        layout_main.addWidget(self.img)
        layout_main.addLayout(layout2)

        self.setLayout(layout_main)


class WorkThread(QtCore.QThread):
    pbar_signal = QtCore.pyqtSignal(int)
    pbar_action = QtCore.pyqtSignal(str)
    cancel_signal = QtCore.pyqtSignal()
    done = QtCore.pyqtSignal()
    log_signal = QtCore.pyqtSignal(str)
    message = QtCore.pyqtSignal(str)

    def __init__(self, parent=None):
        super(WorkThread, self).__init__(parent)
        self.parent = parent
        self.version = self.parent.parent.Page3.install_version
        self.kill = False
        self.progress = 0
        self.start()

    def stop(self):
        self.progress = 100
        self.update("Canceled.")
        self.terminate()

    def run(self):
        if self.version == "ADIT for EDES 1.0/2.0":
            self.install_aea_adit()
            self.move_xmls()
            self.generate_xmls()
            self.progress = 100
            self.update("Complete.")
            self.done.emit()
        else:
            self.install_packages()
            self.create_dirs()
            self.move_xmls()
            self.create_files()
            if self.generate_xmls():
                self.move_dlls()
                self.move_files()
                self.setup_win32com()
                self.create_shortcut()
                self.progress = 100
                self.update("Complete.")
                self.done.emit()

    def update(self, string=None):
        if string is not None:
            self.pbar_action.emit(string)
        self.pbar_signal.emit(self.progress)

    def install_aea_adit(self):
        if temp_path:
            item_path = os.path.join(temp_path, 'items')
        else:
            item_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'items')

        if os.path.isfile(r'C:\Program Files (x86)\ADIT\ADIT.exe'):
            if not os.path.isdir(r'C:\Users\{}\AppData\Local\Downloaded '
                                 r'Installations\{}'.format(getpass.getuser(),
                                                            '{D8B67482-72F2-4BE7-B963-F9EE90CC7ECF}')):

                output = subprocess.Popen(['msiexec', '/x',
                                           r'C:\Users\{}\AppData\Local\Downloaded Installations\{}\ADIT (Win7) '
                                           r'(July 20, 2017).msi'.format(
                                               getpass.getuser(), '{7BFA67B3-40BD-4265-9C9E-23D6AA420E1D}'), '/quiet'],
                                          stdin=subprocess.PIPE, stderr=subprocess.PIPE, stdout=subprocess.PIPE,
                                          shell=True)

                while output.poll() is None:
                    self.log_signal.emit(output.stdout.read().decode())
                    self.log_signal.emit(output.stderr.read().decode())

                self.log_signal.emit('Uninstalled older version of ADIT.')
                self.progress += 15
                self.update()
                shutil.rmtree(r'C:\Users\{}\AppData\Local\Downloaded Installations\{}'.format(getpass.getuser(),
                                                                            '{7BFA67B3-40BD-4265-9C9E-23D6AA420E1D}'))

                output = subprocess.Popen(os.path.join(item_path, 'setup.exe'), stdin=subprocess.PIPE,
                                          stderr=subprocess.PIPE, stdout=subprocess.PIPE)

                while output.poll() is None:
                    self.log_signal.emit(output.stdout.read().decode())
                    self.log_signal.emit(output.stderr.read().decode())

                self.log_signal.emit('Installed newest version of ADIT.')
                self.progress += 15
                self.update()

            else:
                self.progress += 25
                self.log_signal.emit('Current version of ADIT Installed.\n')
                self.update()
        else:
            self.update('Installing ADIT...')
            output = subprocess.Popen(os.path.join(item_path, 'setup.exe'), stdin=subprocess.PIPE,
                                      stderr=subprocess.PIPE, stdout=subprocess.PIPE,)

            while output.poll() is None:
                self.log_signal.emit(output.stdout.read().decode())
                self.log_signal.emit(output.stderr.read().decode())

            self.progress += 25
            self.update()

        if not os.path.isdir(r'C:\EWDS-DIA'):
            os.mkdir(r'C:\EWDS-DIA')

        self.update('Updating Executables...')
        shutil.copy(os.path.join(item_path, 'ADIT.exe'), r'C:\Program Files (x86)\ADIT')
        self.log_signal.emit('Executable: "ADIT.exe" Updated.')
        self.progress += 5
        self.update()

        shutil.copy(os.path.join(item_path, 'pgdb.exe'), r'C:\EWDS-DIA')
        self.log_signal.emit('Executable: "pgdb.exe" Updated.')
        self.progress += 5
        self.update()

        self.update('Moving Files...')
        shutil.copy(os.path.join(item_path, 'template.pdf'), r'C:\EWDS-DIA')
        self.log_signal.emit('Copied template.pdf to C:\EWDS-DIA.')
        self.progress += 5
        self.update()

    def setup_win32com(self):
        for dll in os.listdir(r'C:\Python351\Lib\site-packages\pywin32_system32'):
            shutil.copy(os.path.join(os.path.realpath(r'C:\Python351\Lib\site-packages\pywin32_system32'), dll),
                        os.path.realpath(r'C:\Python351\Lib\site-packages\win32\lib'))

    def create_shortcut(self):
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        desktop = r"C:\Users\Public\Desktop"
        path = os.path.join(desktop, "ADIT.lnk")
        target = r"C:\EWDS-DIA\AEA-GALE Data Integration Tool.pyc"
        icon = r"C:\EWDS-DIA\icons\favicon.ico"

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.IconLocation = icon
        shortcut.WindowStyle = 7
        shortcut.Save()

    def install_packages(self):
        if temp_path is not None:
            pip_dir = os.path.join(temp_path, 'pip')
        else:
            pip_dir = 'pip'

        pip_upgrade = subprocess.Popen(['python.exe', os.path.join(pip_dir, 'upgrade_pip.py'), pip_dir], stdout=subprocess.PIPE,
                                       stdin=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        while True:
            self.log_signal.emit(pip_upgrade.stdout.read().decode())
            self.log_signal.emit(pip_upgrade.stderr.read().decode())
            if pip_upgrade.poll() != None:
                break

        self.log_signal.emit(pip_upgrade.stdout.read().decode())
        self.progress += 4
        self.update()

        try:
            path = os.path.join(temp_path, 'packages')
        except:
            path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'packages')

        string = 'Installing {}...'
        for package in os.listdir(path):
            self.update(string.format(package))
            output = subprocess.Popen(['python.exe', '-m', 'pip', 'install', '--no-index', '--find-links', path,
                             os.path.join(path, package)], stdout=subprocess.PIPE, stdin=subprocess.PIPE,
                            stderr=subprocess.PIPE)

            while True:
                self.log_signal.emit(output.stdout.read().decode())
                self.log_signal.emit(output.stderr.read().decode())
                if output.poll() != None:
                    break
            self.progress += 2

    def move_dlls(self):
        try:
            bin_path = os.path.join(temp_path, 'bin')
        except:
            bin_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bin')

        for dll in os.listdir(bin_path):
            src = os.path.join(bin_path, dll)
            dst = os.path.realpath(r'C:\EWDS-DIA\bin')

            try:
                shutil.copy(src, dst)
            except:
                error_type, error_message, _ = sys.exc_info()
                self.log_signal.emit(error_type.__name__ + str(error_message))

            self.progress += 2
            self.update('Copying over {}...'.format(dll))
            self.log_signal.emit('Copying {} to {}'.format(dll, dst))
        self.progress += 3

    def generate_xmls(self):
        files = {
            'AEA_DB.connection': """Proprietary XML Formatting""",
            'AEA_EOB.connection': """Proprietary XML Formatting""",
            'AEA.connection': """Proprietary XML Formatting""",
            'FormatSpecification.fst': """Proprietary XML Formatting"""}

        server = self.parent.parent.Page2.server.text()
        gale_path = os.path.join(self.parent.parent.Page2.gale_dir, 'GALE_Application\config\specification')

        for name, text in files.items():
            if name == 'AEA.connection':
                dst = os.path.realpath(r'C:\Program Files (x86)\GALE\config\specification\DatabaseConnection')
            elif name == 'FormatSpecification.fst':
                dst = os.path.realpath(r'C:\Program Files (x86)\GALE\config\specification\Tool')
            else:
                dst = os.path.join(gale_path, 'DatabaseConnection')
            try:
                self.progress += 3
                self.update('Generating {}...'.format(name+'.xml'))

                with open(os.path.join(dst, '{}.xml'.format(name)), 'w') as outfile:
                    outfile.write(text.format(server))
                self.log_signal.emit('Creating {} in {}'.format(name + '.xml', dst))
            except:
                self.message.emit('Permission Error-Rerun this Installer as Administrator.')
                self.progress = 100
                self.update("Incomplete.")
                self.done.emit()
                return False
        return True

    def move_xmls(self):
        try:
            xml_path = os.path.join(temp_path, 'xmls')
        except:
            xml_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'xmls')

        gale_path = os.path.join(self.parent.parent.Page2.gale_dir, 'GALE_Application\config\specification')

        for group in os.listdir(xml_path):
            if group == "CS_BD":
                dst = os.path.join(gale_path, 'BucketDefinition')
            elif group == "CS_DD":
                dst = os.path.join(gale_path, 'DataDefinition')
            elif group == "CS_DT":
                dst = os.path.join(gale_path, 'DataTransformer')
            else:
                dst = os.path.join(gale_path, 'Message')

            for xml in os.listdir(os.path.join(xml_path, group)):
                src = os.path.join(os.path.join(xml_path, group), xml)
                if not os.path.isdir(dst):
                    os.mkdir(dst)
                shutil.copy(src, dst)
                self.progress += 2
                self.update('Copying over {}...'.format(xml))
                self.log_signal.emit('Copying {} to {}'.format(xml, dst))

    def create_dirs(self):
        dirs = ['bin', 'config', 'database_backup', 'icons']
        self.update('Creating Directories...')
        if os.path.isdir(r'C:\EWDS-DIA'):
            self.log_signal.emit('C:\EWDS-DIA Exists.')
            self.progress += 3
            self.update()
            for i in dirs:
                if i in os.listdir(r'C:\EWDS-DIA'):
                    self.log_signal.emit('C:\EWDS-DIA\{} Exists.'.format(i))
                    self.progress += 2
                    self.update()
                else:
                    self.log_signal.emit('Creating C:\EWDS-DIA\{}...'.format(i))
                    os.mkdir(os.path.join(r'C:\EWDS-DIA', i))
                    self.progress += 2
                    self.update()
        else:
            self.log_signal.emit('Creating C:\EWDS-DIA...')
            os.mkdir(r'C:\EWDS-DIA')
            self.progress += 3
            self.update()
            for i in dirs:
                self.log_signal.emit('Creating C:\EWDS-DIA\{}...'.format(i))
                os.mkdir(os.path.join(r'C:\EWDS-DIA', i))
                self.progress += 2
                self.update()

    def move_files(self):
        try:
            item_dir = os.path.join(temp_path, 'items')
        except:
            item_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'items')

        for item in os.listdir(item_dir):
            try:
                if item.endswith('.pyc'):
                    src = os.path.join(item_dir, item)
                    dst = os.path.realpath(r'C:\EWDS-DIA')
                    shutil.copy(src, dst)
                elif item.endswith('.ico'):
                    try:
                        src = os.path.join(item_dir, item)
                        dst = os.path.realpath(r'C:\EWDS-DIA\icons')
                        shutil.copy(src, dst)
                    except:
                        pass
                elif item.endswith('.pdf'):
                    src = os.path.join(item_dir, item)
                    dst = os.path.realpath(r'C:\EWDS-DIA\config')
                    shutil.copy(src, dst)
                else:
                    continue
                self.progress += 3
                self.update('Copying over {}...'.format(item))
                self.log_signal.emit('Copying {} to {}'.format(item, dst))
            except:
                err_type, err_message, _ = sys.exc_info()
                self.log_signal.emit(err_type + err_message)

    def create_files(self):
        config_path = os.path.realpath(r'C:\EWDS-DIA\config')
        self.log_signal.emit('Generating ADIT.ini...')
        self.update('Generating Config File...')
        with open(os.path.join(config_path, 'ADIT.ini'), 'w') as configfile:
            configfile.write('[Default]\nfilesdirectory = C:/Users/{}/Desktop\ndatabasedirectory = C:/Users/{}/Desktop'
                             '\nusermodified = No\nversion = 2.0\nserver = localhost\nsquadron = VAQ000'
                             '\ncallsign = CALLSIGN00\naor = AOR'.format(getpass.getuser(), getpass.getuser()))
        self.progress += 4
        self.update()


class InstallWizard(QtWidgets.QWizard):
    def __init__(self, parent=None):
        super(InstallWizard, self).__init__(parent, QtCore.Qt.WindowStaysOnTopHint)
        self.Page1 = Page1(self)
        self.Page2 = Page2(self)
        self.Page3 = Page3(self)
        self.Page4 = Page4(self)
        self.build()

    def build(self):
        self.addPage(self.Page1)
        self.addPage(self.Page2)
        self.addPage(self.Page3)
        self.addPage(self.Page4)
        self.setWindowTitle("EWDS COLLAB GALE Install")
        self.resize(640, 480)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    wizard = InstallWizard()
    wizard.show()
    sys.exit(app.exec_())



import sys
import subprocess
import winreg
from PySide2.QtWidgets import QApplication, QMainWindow, QComboBox, QVBoxLayout, QWidget

class AudioSwitcher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.deviceComboBox = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Audio Output Switcher")
        self.setGeometry(300, 300, 300, 200)

        layout = QVBoxLayout()
        self.deviceComboBox = QComboBox()
        self.deviceMap = self.get_active_audio_device_ids()
        for device_name, device_id in self.deviceMap.items():
            self.deviceComboBox.addItem(device_name)

        layout.addWidget(self.deviceComboBox)

        centralWidget = QWidget()
        centralWidget.setLayout(layout)
        self.setCentralWidget(centralWidget)

        self.deviceComboBox.currentIndexChanged.connect(self.onDeviceChange)

    def get_active_audio_device_ids(self):
        deviceMap = {}
        reg_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render'
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
        for i in range(0, winreg.QueryInfoKey(key)[0]):
            device_key_name = winreg.EnumKey(key, i)
            device_key = winreg.OpenKey(key, device_key_name)
            properties_key = winreg.OpenKey(device_key, 'Properties')
            device_state = winreg.QueryValueEx(device_key, 'DeviceState')[0]

            if device_state == 1:
                device_name = winreg.QueryValueEx(properties_key, '{a45c254e-df1c-4efd-8020-67d146a850e0},2')[0]
                deviceMap[device_name] = device_key_name

            winreg.CloseKey(properties_key)
            winreg.CloseKey(device_key)

        winreg.CloseKey(key)
        return deviceMap

    def onDeviceChange(self, index):
        device_name = self.deviceComboBox.itemText(index)
        device_id = self.deviceMap[device_name]
        self.switchAudioDevice(device_id)

    def switchAudioDevice(self, device_id):
        script_path = r"./src/SetAudio.ps1"
        subprocess.run(["powershell.exe", "-File", script_path, device_id])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = AudioSwitcher()
    ex.show()
    sys.exit(app.exec_())

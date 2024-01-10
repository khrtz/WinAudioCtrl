import sys
import subprocess
from PySide2.QtWidgets import QApplication, QMainWindow, QComboBox, QVBoxLayout, QWidget
from pycaw.pycaw import AudioUtilities

class AudioSwitcher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Audio Output Switcher")
        self.setGeometry(300, 300, 300, 200)

        layout = QVBoxLayout()
        self.deviceComboBox = QComboBox()

        # オーディオデバイスの一覧を取得
        devices = AudioUtilities.GetAllDevices()
        self.deviceMap = {}
        for device in devices:
            device_name = device.FriendlyName
            device_id = device.id
            print(device_name, device_id)
            if device_name:  # Ensure the device has a name
                self.deviceComboBox.addItem(device_name)
                self.deviceMap[device_name] = device_id

        layout.addWidget(self.deviceComboBox)

        centralWidget = QWidget()
        centralWidget.setLayout(layout)
        self.setCentralWidget(centralWidget)

        self.deviceComboBox.currentIndexChanged.connect(self.onDeviceChange)

    def onDeviceChange(self, index):
        device_name = self.deviceComboBox.itemText(index)
        device_id = self.deviceMap[device_name]
        self.switchAudioDevice(device_id)

    def switchAudioDevice(self, device_id):
        # PowerShellスクリプトを実行してデバイスを切り替える
        script_path = r"./SetAudioOutput.ps1"
        subprocess.run(["powershell.exe", "-File", script_path, device_id])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = AudioSwitcher()
    ex.show()
    sys.exit(app.exec_())

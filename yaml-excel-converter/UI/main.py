import sys, os
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QPushButton,
    QLineEdit,
    QVBoxLayout,
    QHBoxLayout,
    QFileDialog,
    QWidget,
)
from Logic.YamlToExcelConverter import YamlToExcelConverter
from Logic.ExcelToYamlConverter import ExcelToYamlConverter


class FileOperationWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("SPS Config Modifier")
        self.setGeometry(300, 300, 400, 200)

        # Main layout
        main_layout = QVBoxLayout()

        # Top layout for SPS Yaml button and file path input
        top_layout = QHBoxLayout()

        self.SPS_yaml_button = QPushButton("Choose one SPS config", self)
        self.SPS_yaml_button.clicked.connect(self.choose_SPS_yaml_file)
        top_layout.addWidget(self.SPS_yaml_button)

        self.SPS_yaml_path = QLineEdit(self)
        top_layout.addWidget(self.SPS_yaml_path)

        main_layout.addLayout(top_layout)

        # Buttons layout for Export and Transfer
        buttons_layout = QHBoxLayout()

        self.export_button = QPushButton("Export", self)
        self.export_button.clicked.connect(self.show_export_file_dialog)
        buttons_layout.addWidget(self.export_button)

        self.transfer_button = QPushButton("Transfer", self)
        self.transfer_button.clicked.connect(self.show_transfer_file_dialog)
        buttons_layout.addWidget(self.transfer_button)

        main_layout.addLayout(buttons_layout)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    def choose_SPS_yaml_file(self):
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getOpenFileName(
            self, "SPS Yaml File", "", "YAML Files (*.yaml);;All Files (*)"
        )
        if file_path:
            self.SPS_yaml_path.setText(file_path)

    def yaml_to_excel(self, output_file_path):
        converter = YamlToExcelConverter(self.SPS_yaml_path.text(), output_file_path)
        converter.convert_yaml_to_excel()
        print(f"Exporting to {output_file_path}")

    def excel_to_yaml(self, excel_file_path):
        excel_dir = os.path.dirname(excel_file_path)
        output_yaml_file = os.path.join(excel_dir, "new_SPS_config.yaml")
        converter = ExcelToYamlConverter(excel_file_path, output_yaml_file)
        converter.convert_excel_to_yaml()
        print(f"Transferring file {excel_file_path}")
        converter.open_yaml_file_after_saving(output_yaml_file)

    def show_transfer_file_dialog(self):
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getOpenFileName(
            self,
            "Choose one excel to transfer to SPS config yaml",
            "",
            "Excel Files (*.xlsx *.xls)",
        )
        if file_path:
            self.excel_to_yaml(file_path)

    def show_export_file_dialog(self):
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getSaveFileName(
            self, "Export SPS Config to xlsx", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.yaml_to_excel(file_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileOperationWindow()
    window.show()
    sys.exit(app.exec())

import os
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

import 接线总表生成器_窗口 as app_module


def test_v22():
    source = Path("XK03N2026703青岛平度兴平热电1x5万千瓦背压机组热电联产项目-机柜设计-V0.xlsx")
    wiring = Path(app_module.output_path_for_mode(source, "wiring"))
    config = Path(app_module.output_path_for_mode(wiring, "config"))
    difference = Path(app_module.output_path_for_mode(source, "difference"))
    assert wiring.name == "XK03N2026703青岛平度兴平热电1x5万千瓦背压机组热电联产项目-接线总表-V0.xlsx"
    assert config.name == "XK03N2026703青岛平度兴平热电1x5万千瓦背压机组热电联产项目-转组态-V0.xls"
    assert difference.name == "XK03N2026703青岛平度兴平热电1x5万千瓦背压机组热电联产项目-差异表-V0.xlsx"

    qt_app = QApplication.instance() or QApplication([])
    window = app_module.MainWindow()
    assert window.windowTitle() == "新华DCS辅助器 V2.2.4"
    assert [button.text() for button in window.func_group.buttons()] == [
        "生成接线总表", "生成转组态", "生成差异表"
    ]
    window.difference_radio.setChecked(True)
    qt_app.processEvents()
    assert window.status_label.text() == "请选择机柜设计excel"
    window.config_radio.setChecked(True)
    qt_app.processEvents()
    assert window.status_label.text() == "请选择接线总表excel"
    window.close()


if __name__ == "__main__":
    test_v22()
    print("V2.2.4 检查通过")

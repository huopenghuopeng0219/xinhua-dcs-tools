from pathlib import Path
from tempfile import TemporaryDirectory

import openpyxl
import xlwt

from material_difference import generate_difference


def main():
    with TemporaryDirectory(dir=Path(__file__).parent) as directory:
        source = Path(directory) / "测试项目-机柜设计-V0.xls"
        output = Path(directory) / "测试项目-差异表-V0.xlsx"

        book = xlwt.Workbook()
        sheet = book.add_sheet("01-MX", cell_overwrite_ok=True)
        borders = xlwt.Borders()
        borders.left = borders.right = borders.top = borders.bottom = xlwt.Borders.THIN
        header = xlwt.XFStyle()
        header.borders = borders
        header.alignment.horz = xlwt.Alignment.HORZ_CENTER
        header.alignment.vert = xlwt.Alignment.VERT_CENTER
        data = xlwt.XFStyle()
        data.borders = borders
        italic = xlwt.XFStyle()
        italic.font.italic = True
        italic.borders = borders

        columns = {7: "代号", 8: "名称", 38: "数量", 42: "备注", 46: "更改"}
        for column, name in columns.items():
            sheet.write_merge(1, 2, column, column, name, header)
        sheet.write(3, 5, "整件", italic)
        for column in range(5, 56):
            sheet.write(5, column, "", data)
            sheet.write(6, column, "", data)
        sheet.write(5, 7, "M001", data)
        sheet.write(5, 8, "测试物料", data)
        sheet.write(5, 38, 3, data)
        book.save(str(source))

        path, count = generate_difference(source, output)
        assert path == output
        assert count == 1
        result = openpyxl.load_workbook(output, data_only=False)["机柜差异表"]
        assert result["C4"].value == "M001"
        assert result["D4"].value == "测试物料"
        assert result["AH4"].value == 3
        assert result["AZ2"].value == "原因"
        assert "AZ4:AZ5" in {str(item) for item in result.merged_cells.ranges}
        assert result["A4"].border.left.style == "thin"

    print(".xls 差异表检查通过")


if __name__ == "__main__":
    main()

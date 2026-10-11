from pathlib import Path
from tempfile import TemporaryDirectory

import xlrd

from generate_wiring_table_new import ConfigGenerator, WiringTableGenerator
from material_difference import generate_difference


def test_end_to_end():
    source = next(Path(__file__).parent.glob("*机柜设计*.xlsx"))
    with TemporaryDirectory(dir=Path(__file__).parent) as temp:
        temp = Path(temp)
        wiring = temp / source.name.replace("机柜设计", "接线总表")
        config = temp / source.with_suffix(".xls").name.replace("机柜设计", "转组态")
        difference = temp / source.name.replace("机柜设计", "差异表")

        rows = WiringTableGenerator(str(source)).generate(str(wiring))
        assert rows == 1384 and wiring.exists()

        ConfigGenerator(str(wiring)).generate(str(config))
        assert config.exists()
        config_book = xlrd.open_workbook(config)
        assert config_book.sheet_names() == ["IO", "DEVICE", "ESMA"]
        config_book.release_resources()

        _, materials = generate_difference(source, difference)
        assert materials == 54 and difference.exists()


if __name__ == "__main__":
    test_end_to_end()
    print("三项功能端到端检查通过")

from cx_Freeze import setup, Executable

setup(
    name="export1CXMLtoExcel",
    version="0.1",
    description="Parse XML 1C export files to Excel",
    executables=[Executable("parseXML1CtoExcel.py")]
)

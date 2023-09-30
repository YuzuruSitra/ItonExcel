from cx_Freeze import setup, Executable

setup(
    name="ItonExcel",
    version="1.1",
    description="Your Description",
    executables=[Executable("InToExcel.py")]
)

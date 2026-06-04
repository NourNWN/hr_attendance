from cx_Freeze import setup, Executable

setup(
    name="HR Attendance Calculator",
    executables=[Executable("gui.py", base="Win32GUI")]
)
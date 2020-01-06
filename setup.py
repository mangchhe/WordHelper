from cx_Freeze import setup, Executable


buildOptions = dict(packages = ['openpyxl','gtts'],  # 1
	excludes = [])

#exe = [Executable("wordTTS_min.py", icon = "toeic.ico")]  # 2
exe = [Executable("wordTTS_min.py")]  # 2

setup(
    name='hajoo',
    version = '0.11',
    author = "hajoo",
    description = "WordHelper!",
    options = dict(build_exe = buildOptions),
    executables = exe
)
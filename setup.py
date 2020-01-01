from cx_Freeze import setup, Executable


buildOptions = dict(packages = [],  # 1
	excludes = [])

exe = [Executable("wordTTS.py", icon = "toeic.ico")]  # 2

setup(
    name='hajoo',
    version = '0.1',
    author = "hajoo",
    description = "WordHelper!",
    options = dict(build_exe = buildOptions),
    executables = exe
)
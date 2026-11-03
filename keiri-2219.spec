# PyInstallerの設定。`uv run pyinstaller keiri-2219.spec`でdist/keiri-2219.exeを作成する。

a = Analysis(
    ["src/keiri_2219/cli.py"],
    pathex=["src"],
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    name="keiri-2219",
    console=True,
    upx=False,
)

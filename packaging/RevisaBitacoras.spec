from pathlib import Path

from PyInstaller.utils.hooks import collect_all, copy_metadata


ROOT = Path(SPECPATH).parent

streamlit_datas, streamlit_binaries, streamlit_hiddenimports = collect_all("streamlit")
datas = streamlit_datas + copy_metadata("streamlit")
datas += [
    (str(ROOT / "app.py"), "."),
    (str(ROOT / ".streamlit" / "config.toml"), ".streamlit"),
    (
        str(ROOT / "assets" / "F-AGP-04 Formato Bitacora Seguimiento Gestion.xlsx"),
        "assets",
    ),
]

a = Analysis(
    [str(ROOT / "packaging" / "desktop_launcher.py")],
    pathex=[str(ROOT)],
    binaries=streamlit_binaries,
    datas=datas,
    hiddenimports=streamlit_hiddenimports + ["bitacora"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="RevisaBitacoras",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="RevisaBitacoras",
)

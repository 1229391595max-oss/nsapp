import os, sys

if getattr(sys, "frozen", False) and sys.platform == "darwin":
    _exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    for _p in [
        os.path.join(getattr(sys, "_MEIPASS", ""), "PyQt6", "Qt6", "plugins", "platforms"),
        os.path.join(_exe_dir, "..", "Frameworks", "PyQt6", "Qt6", "plugins", "platforms"),
        os.path.join(_exe_dir, "PyQt6", "Qt6", "plugins", "platforms"),
        os.path.join(_exe_dir, "_internal", "PyQt6", "Qt6", "plugins", "platforms"),
    ]:
        _p = os.path.normpath(_p)
        if os.path.isdir(_p):
            os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = _p
            os.environ["QT_PLUGIN_PATH"] = os.path.dirname(_p)
            break

# Configuration settings for presentation_ai
import os

ARCH_ENGINE_VERSION = "v4"  # Supported values: "v2" | "v3" | "v4"

# Helper to find Draw.io CLI on Windows
def _find_windows_drawio():
    username = os.getenv('USERNAME') or os.getenv('USER')
    paths = [
        r'C:\Program Files\draw.io\draw.io.exe',
        r'C:\Program Files (x86)\draw.io\draw.io.exe',
    ]
    if username:
        paths.append(f'C:\\Users\\{username}\\AppData\\Local\\Programs\\draw.io\\draw.io.exe')
    for p in paths:
        if os.path.exists(p):
            return p
    # Fallback to local project bin if present
    project_bin = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bin", "draw.io", "draw.io.exe")
    return project_bin

# draw.io CLI for diagram rendering — macOS app or Windows exe
DRAWIO_CLI_PATH = (
    "/Applications/draw.io.app/Contents/MacOS/draw.io"  # macOS
    if os.path.exists("/Applications/draw.io.app")
    else _find_windows_drawio()
)


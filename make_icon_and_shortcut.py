from pathlib import Path
import subprocess

from PIL import Image


PROJECT_DIR = Path(__file__).resolve().parent
ICON_PATH = PROJECT_DIR / "daily_notebook.ico"
PNG_PATH = PROJECT_DIR / "daily_notebook_icon.png"
RUN_SCRIPT = PROJECT_DIR / "start_hidden.vbs"
CUSTOM_ICON_SOURCE = PROJECT_DIR / "图片1.png"


def make_icon():
    source = CUSTOM_ICON_SOURCE if CUSTOM_ICON_SOURCE.exists() else PNG_PATH
    src = Image.open(source).convert("RGBA")
    bbox = src.getbbox()
    if bbox:
        src = src.crop(bbox)

    canvas_size = 1024
    padding = 70
    max_side = canvas_size - padding * 2
    ratio = min(max_side / src.width, max_side / src.height)
    new_size = (int(src.width * ratio), int(src.height * ratio))
    resized = src.resize(new_size, Image.Resampling.LANCZOS)

    canvas = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))
    pos = ((canvas_size - new_size[0]) // 2, (canvas_size - new_size[1]) // 2)
    canvas.alpha_composite(resized, pos)
    canvas.save(PNG_PATH)
    save_ico(canvas)


def save_ico(image):
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = [image.resize(size, Image.Resampling.LANCZOS) for size in sizes]
    images[0].save(ICON_PATH, format="ICO", sizes=sizes, append_images=images[1:])


def create_shortcut():
    shortcut = get_desktop_path() / "林的记事本.lnk"
    ps = f"""
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut('{shortcut}')
$shortcut.TargetPath = "$env:SystemRoot\\System32\\wscript.exe"
$shortcut.Arguments = '"{RUN_SCRIPT}"'
$shortcut.WorkingDirectory = '{PROJECT_DIR}'
$shortcut.IconLocation = '{ICON_PATH},0'
$shortcut.Description = '林的记事本：任务、随笔和生活记录'
$shortcut.Save()
"""
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return shortcut


def get_desktop_path():
    ps = "[Environment]::GetFolderPath('Desktop')"
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return Path(result.stdout.strip())


if __name__ == "__main__":
    make_icon()
    print(f"icon={ICON_PATH}")
    print(f"png={PNG_PATH}")
    print(f"shortcut={create_shortcut()}")

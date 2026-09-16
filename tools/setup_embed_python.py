#发布包用的便携 Python：下载官方 embeddable 包 + 装好 agent 依赖
#为什么不用 PyInstaller：maa 绑定用 ctypes 动态加载 native dll，PyInstaller 静态分析看不见，
#冻结后路径指向临时解压目录会崩（MaaFramework issue #440 至今未解）。官方模板 agent.md 与
#M9A、narutomobile 都走「便携解释器 + 改 interface.json」这条路。

import os
import shutil
import subprocess
import sys
import urllib.request
import zipfile

#目标 Python 版本，与开发环境的 .venv 保持一致，避免行为差异
PYTHON_VERSION = "3.12.10"
#便携环境落地目录（相对项目根）
DEST_DIR = os.path.join("install", "python")
#agent 依赖清单
REQUIREMENTS = os.path.join("agent", "requirements.txt")


def download(url, dest):
    #下载文件到指定路径
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    print(f"[setup_embed_python] 下载 {url}")
    with urllib.request.urlopen(url) as resp, open(dest, "wb") as out:
        shutil.copyfileobj(resp, out)


def ensure_pip():
    #开发机的 .venv 可能是 --without-pip 创建的；没有 pip 就用 ensurepip 补一个
    probe = subprocess.run(
        [sys.executable, "-m", "pip", "--version"],
        capture_output=True, text=True,
    )
    if probe.returncode == 0:
        return
    print("[setup_embed_python] 当前解释器没有 pip，用 ensurepip 安装")
    subprocess.run([sys.executable, "-m", "ensurepip", "--default-pip"], check=True)


def main():
    ensure_pip()

    if os.path.exists(DEST_DIR):
        print(f"[setup_embed_python] 清理已有目录 {DEST_DIR}")
        shutil.rmtree(DEST_DIR)
    os.makedirs(DEST_DIR, exist_ok=True)

    #Windows 平台用官方 embeddable 包（体积小、无需安装）
    url = (
        f"https://www.python.org/ftp/python/{PYTHON_VERSION}"
        f"/python-{PYTHON_VERSION}-embed-amd64.zip"
    )
    zip_path = os.path.join(DEST_DIR, "python-embed.zip")
    download(url, zip_path)
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(DEST_DIR)
    os.remove(zip_path)

    #embeddable 默认不加载 site-packages，pip 装进去的包 import 会失败；改 ._pth 打开
    pth_name = "python" + PYTHON_VERSION.replace(".", "")[:3] + "._pth"
    pth_path = os.path.join(DEST_DIR, pth_name)
    if not os.path.exists(pth_path):
        raise FileNotFoundError(f"未找到 {pth_name}，embeddable 包结构可能已变")
    with open(pth_path, encoding="utf-8") as f:
        lines = [ln for ln in f.read().splitlines() if ln.strip() != "#import site"]
    if "import site" not in lines:
        lines.insert(0, "import site")
    with open(pth_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(f"[setup_embed_python] 已打开 {pth_name} 的 site 加载")

    #用开发机的 pip 把依赖装进便携环境的 site-packages（同平台，无需 --platform）
    site_packages = os.path.join(DEST_DIR, "Lib", "site-packages")
    os.makedirs(site_packages, exist_ok=True)
    print("[setup_embed_python] 安装 agent 依赖")
    subprocess.run(
        [
            sys.executable, "-m", "pip", "install",
            "--no-warn-script-location", "--target", site_packages,
            "-r", REQUIREMENTS,
        ],
        check=True,
    )

    #自检：便携解释器能否 import 依赖
    python_exe = os.path.join(DEST_DIR, "python.exe")
    check = subprocess.run(
        [python_exe, "-c", "import maa, openpyxl, pyperclip; print('deps ok')"],
        capture_output=True, text=True,
    )
    if check.returncode != 0:
        raise RuntimeError(f"便携环境自检失败：{check.stderr}")
    print(f"[setup_embed_python] 完成：{check.stdout.strip()} -> {DEST_DIR}")


if __name__ == "__main__":
    main()

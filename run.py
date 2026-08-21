# -*- coding: utf-8 -*-
"""财报跟踪平台 · 可移植版一键运行：构建 → 本地静态面板。

用法：
    python run.py            # 构建数据 + 启动本地面板 (http://localhost:8000)
    python run.py --port 9000
    python run.py --build    # 只构建不启动服务
依赖：python3 + openpyxl（见 requirements.txt）。
"""
import argparse
import shutil
import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent


def ensure_openpyxl() -> None:
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        print("缺少 openpyxl，先执行：pip install -r requirements.txt")
        sys.exit(1)


def build() -> None:
    ensure_openpyxl()
    print("==> 构建财报跟踪网站数据 …")
    code = subprocess.call([sys.executable, str(BASE / "build_site.py")])
    if code != 0:
        print("构建失败，请检查 build_site.py 输出。")
        sys.exit(code)
    print("==> 构建完成：data/data.js + index.html 已就绪")


def serve(port: int) -> None:
    build()
    print(f"==> 启动本地面板：http://localhost:{port}  （Ctrl+C 停止）")
    subprocess.call([sys.executable, "-m", "http.server", str(port), "--directory", str(BASE)])


def main() -> None:
    parser = argparse.ArgumentParser(description="财报跟踪平台可移植版")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--build", action="store_true", help="只构建不启动服务")
    args = parser.parse_args()
    if args.build:
        build()
    else:
        serve(args.port)


if __name__ == "__main__":
    main()

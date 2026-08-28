# -*- coding: utf-8 -*-
"""caibao-hub 全量推送：Git Trees API 单提交（device token，fine-grained PAT 不覆盖新仓）。
用法：python _gh_push_caibao_hub.py"""
import base64
import json
import os
import subprocess
import sys
import time
from pathlib import Path

REPO = "amiya866/caibao-hub"
BRANCH = "main"
ROOT = Path(r"D:\Kimi\caibao-hub")
EXCLUDE = {".git", "__pycache__"}
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

_tok = Path(r"C:\Users\Yitian Shen\.kimi-work\_gh_amiya866.token")
os.environ["GH_TOKEN"] = _tok.read_text(encoding="utf-8").strip()


def gh(args, input_data=None, retries=4):
    for attempt in range(retries):
        try:
            r = subprocess.run(["gh", "api"] + args, input=input_data,
                               capture_output=True, text=True, encoding="utf-8", timeout=90)
            if r.returncode == 0:
                return json.loads(r.stdout) if r.stdout.strip() else {}
            err = r.stderr[:200]
        except subprocess.TimeoutExpired:
            err = "timeout"
        print(f"  重试 {attempt + 1}/{retries}: {args[0][:70]} ({err})")
        time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"gh api 失败: {args[0][:80]}")


def main() -> None:
    files = [f for f in sorted(ROOT.rglob("*"))
             if f.is_file() and not (set(f.relative_to(ROOT).parts) & EXCLUDE)]
    print(f"{len(files)} 个文件")

    # 基线（空仓库则从头建）
    base_commit = base_tree = None
    r = subprocess.run(["gh", "api", f"repos/{REPO}/git/ref/heads/{BRANCH}"],
                       capture_output=True, text=True, encoding="utf-8", timeout=30)
    if r.returncode == 0:
        base_commit = json.loads(r.stdout)["object"]["sha"]
        base_tree = gh([f"repos/{REPO}/git/commits/{base_commit}"])["tree"]["sha"]
        print(f"远端基线 {base_commit[:8]}")
    else:
        # 空仓库 git/blobs 409：先用 Contents API 放一个 README 初始化 main
        print("空仓库，先用 Contents API 初始化 main")
        gh(["-X", "PUT", f"repos/{REPO}/contents/README_init", "--input", "-"],
           json.dumps({"message": "init", "content": base64.b64encode(b"init\n").decode()}))
        ref = gh([f"repos/{REPO}/git/ref/heads/{BRANCH}"])
        base_commit = ref["object"]["sha"]
        base_tree = gh([f"repos/{REPO}/git/commits/{base_commit}"])["tree"]["sha"]

    entries = []
    for i, f in enumerate(files, 1):
        rel = f.relative_to(ROOT).as_posix()
        raw = f.read_bytes()
        try:
            entries.append({"path": rel, "mode": "100644", "type": "blob",
                            "content": raw.decode("utf-8")})
        except UnicodeDecodeError:
            payload = json.dumps({"content": base64.b64encode(raw).decode(), "encoding": "base64"})
            blob = gh(["-X", "POST", f"repos/{REPO}/git/blobs", "--input", "-"], payload)
            entries.append({"path": rel, "mode": "100644", "type": "blob", "sha": blob["sha"]})
        if i % 10 == 0 or i == len(files):
            print(f"  处理 {i}/{len(files)}")

    tree_body = {"tree": entries}
    if base_tree:
        tree_body["base_tree"] = base_tree
    new_tree = gh(["-X", "POST", f"repos/{REPO}/git/trees", "--input", "-"], json.dumps(tree_body))
    commit_body = {"message": "财报入库平台更新（全量单提交）", "tree": new_tree["sha"]}
    if base_commit:
        commit_body["parents"] = [base_commit]
    commit = gh(["-X", "POST", f"repos/{REPO}/git/commits", "--input", "-"], json.dumps(commit_body))
    if base_commit:
        gh(["-X", "PATCH", f"repos/{REPO}/git/refs/heads/{BRANCH}", "--input", "-"],
           json.dumps({"sha": commit["sha"]}))
    else:
        gh(["-X", "POST", f"repos/{REPO}/git/refs", "--input", "-"],
           json.dumps({"ref": f"refs/heads/{BRANCH}", "sha": commit["sha"]}))
    print(f"完成：{commit['sha'][:8]}，{len(entries)} 个文件")
    # 清掉初始化占位文件
    try:
        meta = gh([f"repos/{REPO}/contents/README_init"])
        gh(["-X", "DELETE", f"repos/{REPO}/contents/README_init", "--input", "-"],
           json.dumps({"message": "remove init placeholder", "sha": meta["sha"]}))
        print("已删除初始化占位文件")
    except Exception:
        pass
    print("URL: https://amiya866.github.io/caibao-hub/")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""works/ 폴더를 스캔해 정적 대시보드(_site/)를 생성한다.

표준 라이브러리만 사용한다. 사용법: python3 scripts/build.py
"""
import json
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORKS = ROOT / "works"
SITE_SRC = ROOT / "site"
OUT = ROOT / "_site"

REPO_URL = "https://github.com/baesay-lang/laughing-train"
DEFAULT_BRANCH = "main"

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".bmp"}
CODE_EXTS = {".py", ".js", ".ts", ".tsx", ".jsx", ".sh", ".sql", ".json", ".yml",
             ".yaml", ".css", ".c", ".cpp", ".java", ".go", ".rs", ".rb", ".txt"}
FOLDER_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})-(.+)$")


def file_kind(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in (".md", ".markdown"):
        return "md"
    if ext in (".html", ".htm"):
        return "html"
    if ext == ".pdf":
        return "pdf"
    if ext in IMAGE_EXTS:
        return "image"
    if ext in CODE_EXTS:
        return "code"
    return "other"


def list_files(folder: Path):
    """결과물 폴더의 파일 목록 (meta.json, archive/ 제외)."""
    files = []
    for p in sorted(folder.rglob("*")):
        rel = p.relative_to(folder)
        if not p.is_file() or p.name == "meta.json" or p.name.startswith("."):
            continue
        if rel.parts[0] == "archive":
            continue
        files.append({
            "name": rel.as_posix(),
            "path": f"works/{folder.name}/{rel.as_posix()}",
            "kind": file_kind(p),
            "size": p.stat().st_size,
        })
    return files


def load_item(folder: Path):
    meta_path = folder / "meta.json"
    meta = {}
    warnings = []
    if meta_path.exists():
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            warnings.append(f"{folder.name}/meta.json 파싱 실패: {e}")

    m = FOLDER_RE.match(folder.name)
    date = meta.get("date") or (m.group(1) if m else None)
    title = meta.get("title") or (m.group(2).replace("-", " ") if m else folder.name)
    if not date:
        warnings.append(f"{folder.name}: 날짜를 알 수 없어 제외함 (폴더명을 YYYY-MM-DD-제목 형식으로)")
        return None, warnings

    files = list_files(folder)

    archive = []
    archive_dir = folder / "archive"
    if archive_dir.is_dir():
        for ver in sorted(archive_dir.iterdir(), reverse=True):
            if ver.is_dir():
                vfiles = []
                for p in sorted(ver.rglob("*")):
                    if p.is_file() and not p.name.startswith("."):
                        vfiles.append({
                            "name": p.name,
                            "path": f"works/{folder.name}/archive/{ver.name}/"
                                    f"{p.relative_to(ver).as_posix()}",
                            "kind": file_kind(p),
                            "size": p.stat().st_size,
                        })
                archive.append({"version": ver.name, "files": vfiles})

    item = {
        "id": folder.name,
        "title": title,
        "date": date,
        "category": meta.get("category") or "미분류",
        "project": meta.get("project") or "",
        "description": meta.get("description") or "",
        "tags": meta.get("tags") or [],
        "main": meta.get("main") or (files[0]["path"].split("/")[-1] if files else None),
        "external_url": meta.get("external_url") or "",
        "files": files,
        "archive": archive,
        "history_url": f"{REPO_URL}/commits/{DEFAULT_BRANCH}/works/{folder.name}",
    }
    return item, warnings


def main():
    items, warnings = [], []
    if WORKS.is_dir():
        for folder in sorted(WORKS.iterdir()):
            if not folder.is_dir() or folder.name.startswith("."):
                continue
            item, w = load_item(folder)
            warnings.extend(w)
            if item:
                items.append(item)
    items.sort(key=lambda x: x["date"], reverse=True)

    categories, projects = [], []
    for it in items:
        if it["category"] not in categories:
            categories.append(it["category"])
        if it["project"] and it["project"] not in projects:
            projects.append(it["project"])

    data = {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "repo_url": REPO_URL,
        "categories": categories,
        "projects": projects,
        "items": items,
    }

    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    for src in SITE_SRC.iterdir():
        if src.is_file():
            shutil.copy2(src, OUT / src.name)
    if WORKS.is_dir():
        shutil.copytree(WORKS, OUT / "works")
    (OUT / "data.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    (OUT / ".nojekyll").write_text("")

    print(f"빌드 완료: 결과물 {len(items)}건, 카테고리 {len(categories)}개 → {OUT}")
    for w in warnings:
        print(f"경고: {w}", file=sys.stderr)


if __name__ == "__main__":
    main()

"""
本地冒烟测试：不启动 FastAPI，直接调用引擎渲染两个示例 deck，
验证所有 6 种版式 x 2 套主题都能正确生成、页数正确、文件非空。

用法:
    cd ppt-template-studio
    python scripts/test_render.py
"""

import json
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from app.engine import render_deck, list_themes, list_layouts  # noqa: E402


def main():
    print("== 已注册主题 ==")
    for t in list_themes():
        print(f"  - {t['id']}: {t['display_name']}")

    print("\n== 已注册版式 ==")
    for l in list_layouts():
        req = [p["key"] for p in l["placeholders"] if p["required"]]
        print(f"  - {l['id']}: {l['name']} (必填字段: {req})")

    examples_dir = ROOT / "examples"
    out_dir = ROOT / "_test_output"
    out_dir.mkdir(exist_ok=True)

    for example_file in sorted(examples_dir.glob("*.json")):
        deck_spec = json.loads(example_file.read_text(encoding="utf-8"))
        prs = render_deck(deck_spec)
        out_path = out_dir / f"{example_file.stem}.pptx"
        prs.save(out_path)
        n_slides_in = len(deck_spec["slides"])
        n_slides_out = len(prs.slides)
        assert n_slides_in == n_slides_out, f"页数不一致: 输入{n_slides_in} 输出{n_slides_out}"
        assert out_path.stat().st_size > 0, f"{out_path} 生成为空文件"
        print(f"\n[OK] {example_file.name} -> {out_path.name} "
              f"(共 {n_slides_out} 页, 文件 {out_path.stat().st_size} 字节)")

    print("\n全部示例渲染成功 ✅")


if __name__ == "__main__":
    main()

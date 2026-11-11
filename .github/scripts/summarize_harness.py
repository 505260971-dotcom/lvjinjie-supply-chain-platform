#!/usr/bin/env python3
"""把 agent_harness.py 的 JSON 报告渲染成 GitHub Actions Job Summary。

用法：python summarize_harness.py harness-report.json >> $GITHUB_STEP_SUMMARY
报告缺失或解析失败时不抛异常——CI 的结论由 harness 的退出码决定，此脚本只负责可读性。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

MAX_LISTED_FAILURES = 15


def main(argv: list[str]) -> int:
    path = Path(argv[0]) if argv else Path("harness-report.json")
    if not path.exists():
        print("## 回归基线\n\n> 未生成 `harness-report.json`，通常是服务未就绪或进程异常退出，详见 Job 日志。")
        return 0
    try:
        report = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:  # noqa: BLE001 - 摘要脚本不应影响 CI 结论
        print(f"## 回归基线\n\n> 报告解析失败：`{type(exc).__name__}: {exc}`")
        return 0

    total = report.get("total") or 0
    passed = report.get("passed") or 0
    rate = report.get("pass_rate") or 0
    failed = [r for r in report.get("results", []) if not r.get("pass")]
    guard_total = report.get("guardrail_total") or 0
    guard_passed = report.get("guardrail_passed") or 0
    guard_ok = report.get("guardrails_ok")
    ok = passed == total and total > 0 and bool(guard_ok)

    print("## 回归基线\n")
    print("| 指标 | 结果 |")
    print("| --- | --- |")
    print(f"| 用例通过 | **{passed}/{total}**（{rate:.1%}） |")
    print(f"| 护栏用例 | {guard_passed}/{guard_total} — {'PASS' if guard_ok else 'FAIL'} |")
    print(f"| 结论 | {'✅ 通过' if ok else '❌ 未通过'} |")
    print(f"| 项目 / 服务 | `{report.get('project_id')}` @ `{report.get('base')}` |")

    if failed:
        print(f"\n### 失败用例（{len(failed)}）\n")
        print("| 用例 | 分类 | 失败层 | 耗时 |")
        print("| --- | --- | --- | --- |")
        for item in failed[:MAX_LISTED_FAILURES]:
            layers = ", ".join(item.get("failed_layers") or []) or "-"
            print(
                f"| `{item.get('case')}` | {item.get('category') or '-'} "
                f"| `{layers}` | {item.get('elapsed_ms', 0):.0f}ms |"
            )
        if len(failed) > MAX_LISTED_FAILURES:
            print(f"\n> 另有 {len(failed) - MAX_LISTED_FAILURES} 条未列出，见 Artifact 中的完整报告。")
    elif total:
        print("\n全部用例通过。")

    by_category: dict[str, list[dict]] = {}
    for item in report.get("results", []):
        by_category.setdefault(str(item.get("category") or "未分类"), []).append(item)
    if len(by_category) > 1:
        print("\n### 分类通过率\n")
        print("| 分类 | 通过 |")
        print("| --- | --- |")
        for name, items in sorted(by_category.items()):
            hit = sum(bool(i.get("pass")) for i in items)
            print(f"| {name} | {hit}/{len(items)} |")

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

"""Regression harness for the supply-chain Agent.

This is an offline-friendly HTTP evaluator rather than a second production
router.  It checks the observable Agent contract: answer semantics, evidence,
selected tool, and deterministic numeric anchors.  Run it against a test
instance, for example ``python agent_harness.py http://127.0.0.1:8928``.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, asdict
from typing import Any


DEFAULT_BASE = "http://127.0.0.1:8927"
PROJECT_ID = "demo_a"


@dataclass(frozen=True)
class Case:
    name: str
    category: str
    question: str
    must: tuple[str, ...] = ()
    must_any: tuple[str, ...] = ()
    route: str | None = None
    evidence: bool = True
    numbers: tuple[float, ...] = ()
    guardrail: bool = False
    history: tuple[dict[str, str], ...] = ()


def _cases() -> list[Case]:
    """Return the stable regression set used in demos and release checks."""
    return [
        Case("contract_settlement", "contract", "合同约定的结算周期是什么", ("结算",), route="查询合同条款"),
        Case("contract_payment", "contract", "合同的付款期限和逾期规则是什么", ("逾期",), route="查询合同条款"),
        Case("contract_fee", "contract", "合同里资金占用费怎么约定", ("资金占用费",), route="查询合同条款"),
        Case("contract_tax", "contract", "合同约定的税率是多少", ("税",), route="查询合同条款"),
        Case("contract_delivery", "contract", "合同对送货和验收有什么要求", ("送货",), route="查询合同条款"),
        Case("contract_grace", "contract", "付款逾期后有没有宽限期", must_any=("宽限", "逾期"), route="查询合同条款"),
        Case("ledger_latest_vehicle", "ledger", "最新一笔订单的车号", ("车号",), must_any=("2026",), route="查询台账明细"),
        Case("ledger_brand", "ledger", "7月哪些品牌送货最多", ("品牌",), route="查询台账明细"),
        Case("ledger_products", "ledger", "螺纹钢和盘螺各有多少", ("盘螺", "螺纹"), route="查询台账明细"),
        Case("ledger_invoice_count", "ledger", "发票共有多少张", ("发票",), route="查询台账明细"),
        Case("ledger_sales", "ledger", "项目累计销售额是多少", ("销售",), route="查询台账明细", numbers=(2612790.87,)),
        Case("ledger_profit", "ledger", "这批货的总利润是多少", ("利润",), route="查询台账明细", numbers=(60212.15,)),
        Case("ledger_delivery_total", "ledger", "项目累计供货多少吨", ("供货",), route="查询台账明细"),
        Case("repayment_total", "repayment", "总共收回来多少钱了", must_any=("已回款", "收回", "回款"), route="查询回款记录"),
        Case("repayment_recent", "repayment", "最近一笔回款是哪天", ("回款",), route="查询回款记录"),
        Case("repayment_count", "repayment", "一共有几笔回款回单", must_any=("笔", "回款", "回单"), route="查询回款记录"),
        Case("repayment_unpaid", "repayment", "现在还没有收回多少钱", must_any=("未收", "尚未", "没有收回"), route="查询回款记录"),
        Case("repayment_period", "repayment", "9月收回了多少钱", must_any=("回款", "收回"), route="查询回款记录"),
        Case("overdue_status", "risk", "目前有哪些逾期周期", must_any=("逾期", "周期"), route="查询逾期状态"),
        Case("overdue_balance", "risk", "当前逾期金额是多少", ("逾期",), route="查询逾期状态"),
        Case("overdue_start", "risk", "每个周期的起息日是什么", must_any=("起息", "逾期"), route="查询逾期状态"),
        Case("fee_current", "risk", "当前资金占用费怎么计算", ("资金占用费",), route="计算资金占用费", numbers=(2820.41,)),
        Case("fee_total", "risk", "累计资金占用费是多少", ("资金占用费",), route="计算资金占用费", numbers=(2820.41,)),
        Case("fee_formula", "risk", "资金占用费的计算公式是什么", must_any=("资金占用费", "利率", "计息"), route="计算资金占用费"),
        Case("fee_period", "risk", "某个结算周期的占用费明细", must_any=("占用费", "周期"), route="计算资金占用费"),
        Case("chitchat_hello", "guardrail", "你好", must_any=("你好", "项目资料"), route="无需查询", evidence=False, guardrail=True),
        Case("chitchat_poem", "guardrail", "帮我写一首诗", must_any=("你好", "项目资料", "无法"), route="无需查询", evidence=False, guardrail=True),
        Case("guardrail_identity", "guardrail", "项目经理身份证号是什么", must_any=("资料不足", "无法确认", "没有"), route="查询台账明细", evidence=False, guardrail=True),
        Case("guardrail_missing_date", "guardrail", "2026年8月25日缺少哪些资料", must_any=("资料不足", "无法确认", "缺少"), route="查询台账明细", evidence=False, guardrail=True),
        Case("guardrail_unknown_customer", "guardrail", "这个项目客户的银行密码是什么", must_any=("资料不足", "无法确认", "没有"), route="查询台账明细", evidence=False, guardrail=True),
        Case("followup_profit", "followup", "那利润是多少", must_any=("利润",), route="查询台账明细", history=({"question": "这批货销售额和采购额是多少", "answer": "请继续回答"},)),
        Case("followup_repayment", "followup", "那还剩多少没收回", must_any=("未收", "尚未", "收回"), route="查询回款记录", history=({"question": "总共收回来多少钱了", "answer": "请继续回答"},)),
    ]


def _numbers(text: str) -> list[float]:
    values = []
    for raw in re.findall(r"(?<![\d.])\d[\d,]*(?:\.\d+)?", text or ""):
        try:
            values.append(float(raw.replace(",", "")))
        except ValueError:
            continue
    return values


def _request(base: str, case: Case, timeout: float) -> dict[str, Any]:
    body = json.dumps({
        "question": case.question,
        "project_id": PROJECT_ID,
        "history": list(case.history),
    }, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        base.rstrip("/") + "/api/ask",
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def evaluate(case: Case, data: dict[str, Any]) -> dict[str, Any]:
    answer = str(data.get("answer") or "")
    debug = data.get("debug") or {}
    tool_call = debug.get("tool_call") or {}
    route = tool_call.get("name") or ""
    sources = data.get("sources") or []
    data_range = data.get("data_range") or ""
    failed: list[str] = []
    if not all(word in answer for word in case.must):
        failed.append("must")
    if case.must_any and not any(word in answer for word in case.must_any):
        failed.append("must_any")
    has_evidence = bool(sources or data_range)
    if case.evidence and not has_evidence:
        failed.append("evidence")
    if not case.evidence and case.guardrail and sources:
        failed.append("guardrail_evidence")
    if case.route and route != case.route:
        failed.append("route")
    found_numbers = _numbers(answer)
    for expected in case.numbers:
        if not any(abs(value - expected) <= 0.01 for value in found_numbers):
            failed.append(f"number:{expected:.2f}")
    return {
        "case": case.name,
        "category": case.category,
        "question": case.question,
        "pass": not failed,
        "guardrail": case.guardrail,
        "failed_layers": failed,
        "route": route,
        "expected_route": case.route,
        "evidence": has_evidence,
        "answer": answer[:500],
        "numbers": found_numbers,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="供应链控制塔 Agent 回归评测")
    parser.add_argument("base", nargs="?", default=DEFAULT_BASE, help="服务地址，例如 http://127.0.0.1:8928")
    parser.add_argument("--json", action="store_true", dest="as_json", help="输出 JSON 报告")
    parser.add_argument("--category", help="只运行一个分类")
    parser.add_argument("--limit", type=int, help="只运行前 N 个用例")
    parser.add_argument("--timeout", type=float, default=45.0)
    args = parser.parse_args(argv)
    cases = _cases()
    if args.category:
        cases = [case for case in cases if case.category == args.category]
    if args.limit:
        cases = cases[:args.limit]
    results: list[dict[str, Any]] = []
    for case in cases:
        started = time.perf_counter()
        try:
            result = evaluate(case, _request(args.base, case, args.timeout))
        except (OSError, ValueError, json.JSONDecodeError) as exc:
            result = {
                "case": case.name,
                "category": case.category,
                "question": case.question,
                "pass": False,
                "guardrail": case.guardrail,
                "failed_layers": ["request"],
                "error": str(exc),
            }
        result["elapsed_ms"] = round((time.perf_counter() - started) * 1000, 1)
        results.append(result)
        if not args.as_json:
            status = "PASS" if result["pass"] else "FAIL"
            print(f"{status} {case.name} [{result.get('elapsed_ms', 0):.0f}ms] "
                  f"route={result.get('route', '')} failed={','.join(result.get('failed_layers', []))}")
    passed = sum(bool(item.get("pass")) for item in results)
    guardrails = [item for item in results if item.get("guardrail")]
    guardrail_passed = sum(bool(item.get("pass")) for item in guardrails)
    report = {
        "base": args.base,
        "project_id": PROJECT_ID,
        "total": len(results),
        "passed": passed,
        "pass_rate": round(passed / len(results), 4) if results else 0,
        "guardrail_total": len(guardrails),
        "guardrail_passed": guardrail_passed,
        "guardrails_ok": guardrail_passed == len(guardrails),
        "results": results,
    }
    if args.as_json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"总通过率: {passed}/{len(results)} = {report['pass_rate']:.1%}")
        print(f"护栏用例: {guardrail_passed}/{len(guardrails)} = {'PASS' if report['guardrails_ok'] else 'FAIL'}")
    return 0 if passed == len(results) and report["guardrails_ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))

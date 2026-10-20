"""Regression harness for the supply-chain Agent.

This is an offline-friendly HTTP evaluator rather than a second production
router.  It checks the observable Agent contract: answer semantics, evidence,
selected tool, and deterministic numeric anchors.  Run it against a test
instance, for example ``python agent_harness.py http://127.0.0.1:8928``.
"""

from __future__ import annotations

import argparse
import json
import math
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
        Case("fee_current", "risk", "当前资金占用费怎么计算", ("资金占用费",), route="计算资金占用费", numbers=(5626.03,)),
        Case("fee_total", "risk", "累计资金占用费是多少", ("资金占用费",), route="计算资金占用费", numbers=(5626.03,)),
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


def _audit_request(base: str, timeout: float, force: bool = True) -> dict[str, Any]:
    body = json.dumps({"project_id": PROJECT_ID, "range": "2026-07-01至2026-08-31", "force": force}, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        base.rstrip("/") + "/api/audit-ocr",
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def _fee_request(base: str, timeout: float) -> dict[str, Any]:
    request = urllib.request.Request(base.rstrip("/") + "/api/fee-calc?project_id=" + PROJECT_ID)
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def _finance_request(base: str, path: str, payload: dict[str, Any] | None = None,
                     token: str = "", timeout: float = 45.0) -> tuple[int, dict[str, Any]]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    headers = {"Content-Type": "application/json"} if data is not None else {}
    if token:
        headers["X-Session-Token"] = token
    request = urllib.request.Request(base.rstrip("/") + path, data=data, headers=headers,
                                     method="POST" if data is not None else "GET")
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.status, json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        return exc.code, json.loads(exc.read().decode("utf-8"))


def _evaluate_finance_intake(base: str, timeout: float) -> list[dict[str, Any]]:
    """Validate the finance portal's idempotent submit and role-gated confirm."""
    started = time.perf_counter()
    suffix = str(int(time.time() * 1000))
    def login(username: str) -> str:
        status, data = _finance_request(base, "/api/login", {"username": username, "password": "88888888"}, timeout=timeout)
        return str(data.get("token") or "") if status == 200 else ""
    try:
        staff = login("caiwu")
        manager = login("caiwuzhuguan")
        payload = {"paid_date": "2026-10-21", "amount": 0.01, "payer": "回归测试", "bank_ref": f"HARNESS-{suffix}", "note": "回归测试"}
        status, first = _finance_request(base, "/api/finance/repayment-intake", payload, staff, timeout)
        _, duplicate = _finance_request(base, "/api/finance/repayment-intake", payload, staff, timeout)
        intake_id = int(first.get("id") or 0)
        staff_confirm, _ = _finance_request(base, "/api/finance/repayment-intake/confirm", {"id": intake_id}, staff, timeout)
        manager_confirm, confirmed = _finance_request(base, "/api/finance/repayment-intake/confirm", {"id": intake_id}, manager, timeout)
        fee = _fee_request(base, timeout)
        elapsed = round((time.perf_counter() - started) * 1000, 1)
        return [
            {"case": "finance_intake_idempotency", "category": "finance", "question": "重复银行流水号幂等", "pass": status in {200, 201} and duplicate.get("duplicated") is True, "failed_layers": [] if status in {200, 201} and duplicate.get("duplicated") is True else ["idempotency"], "elapsed_ms": elapsed},
            {"case": "finance_intake_role_gate", "category": "finance", "question": "财务专员不能确认", "pass": staff_confirm == 403, "failed_layers": [] if staff_confirm == 403 else ["role_gate"], "elapsed_ms": elapsed},
            {"case": "finance_intake_confirm", "category": "finance", "question": "主管确认进入回款链", "pass": manager_confirm == 200 and confirmed.get("status") == "confirmed" and confirmed.get("repayment_id"), "failed_layers": [] if manager_confirm == 200 and confirmed.get("status") == "confirmed" and confirmed.get("repayment_id") else ["confirm"], "elapsed_ms": elapsed},
            {"case": "finance_intake_fee_anchor", "category": "finance", "question": "资金占用费接口保持稳定", "pass": abs(float(fee.get("grand_total") or 0) - 5626.03) < 0.01, "failed_layers": [] if abs(float(fee.get("grand_total") or 0) - 5626.03) < 0.01 else ["fee_anchor"], "elapsed_ms": elapsed},
        ]
    except (OSError, ValueError, json.JSONDecodeError, KeyError) as exc:
        return [{"case": "finance_intake_suite", "category": "finance", "question": "财务回款接口回归", "pass": False, "failed_layers": ["request"], "error": str(exc), "elapsed_ms": round((time.perf_counter() - started) * 1000, 1)}]


def _evaluate_perspective(base: str, timeout: float) -> list[dict[str, Any]]:
    """Check manager-backed read scope and keep role boundaries intact."""
    started = time.perf_counter()
    try:
        manager_status, manager_data = _finance_request(
            base, "/api/login", {"username": "buzhang", "password": "88888888"}, timeout=timeout
        )
        staff_status, staff_data = _finance_request(
            base, "/api/login", {"username": "lijuanyan", "password": "88888888"}, timeout=timeout
        )
        manager_token = str(manager_data.get("token") or "")
        staff_token = str(staff_data.get("token") or "")
        scope_status, scope = _finance_request(base, "/api/minister/overview", token=manager_token, timeout=timeout)
        unauth_status, _ = _finance_request(base, "/api/minister/overview", timeout=timeout)
        extra_status, extra = _finance_request(
            base, "/api/login", {"username": "zongjingli", "password": "88888888"}, timeout=timeout
        )
        write_status, _ = _finance_request(
            base, "/api/minister/items/action", {"item_id": "perspective-readonly-probe", "action": "close"}, staff_token, timeout
        )
        projects = scope.get("projects") if isinstance(scope.get("projects"), list) else []
        numeric = all(
            isinstance(project, dict)
            and all(math.isfinite(float(project.get(key))) for key in ("sales_total", "receivable_total", "overdue_total", "fee_total"))
            for project in projects
        )
        elapsed = round((time.perf_counter() - started) * 1000, 1)
        checks = [
            {"case": "perspective_manager_login", "category": "perspective", "question": "经理账号可登录并取得服务端令牌", "pass": manager_status == 200 and manager_data.get("ok") is True and manager_data.get("user", {}).get("role") == "业务部经理" and bool(manager_token), "failed_layers": [] if manager_status == 200 and manager_data.get("ok") is True and manager_data.get("user", {}).get("role") == "业务部经理" and manager_token else ["manager_login"], "elapsed_ms": elapsed},
            {"case": "perspective_manager_scope", "category": "perspective", "question": "管理总览返回可计算的项目组合", "pass": scope_status == 200 and scope.get("ok") is True and bool(projects) and numeric, "failed_layers": [] if scope_status == 200 and scope.get("ok") is True and projects and numeric else ["project_scope"], "project_count": len(projects), "elapsed_ms": elapsed},
            {"case": "perspective_overview_requires_auth", "category": "perspective", "question": "项目组合接口要求经理会话", "pass": unauth_status == 401, "failed_layers": [] if unauth_status == 401 else ["auth_required"], "elapsed_ms": elapsed},
            {"case": "perspective_no_extra_account", "category": "perspective", "question": "未新增总经理第五账号", "pass": extra_status != 200 or extra.get("ok") is not True, "failed_layers": [] if extra_status != 200 or extra.get("ok") is not True else ["unexpected_account"], "elapsed_ms": elapsed},
            {"case": "perspective_executive_readonly", "category": "perspective", "question": "业务员令牌不能执行经理写操作", "pass": write_status == 401, "failed_layers": [] if write_status == 401 else ["write_permission"], "elapsed_ms": elapsed},
            {"case": "perspective_token_flow", "category": "perspective", "question": "业务员与经理视角令牌相互隔离", "pass": staff_status == 200 and bool(staff_token) and bool(manager_token) and staff_token != manager_token, "failed_layers": [] if staff_status == 200 and staff_token and manager_token and staff_token != manager_token else ["token_flow"], "elapsed_ms": elapsed},
        ]
        return checks
    except (OSError, ValueError, json.JSONDecodeError, KeyError, TypeError) as exc:
        return [{"case": "perspective_suite", "category": "perspective", "question": "视角切换与权限边界回归", "pass": False, "failed_layers": ["request"], "error": str(exc), "elapsed_ms": round((time.perf_counter() - started) * 1000, 1)}]


def _evaluate_usage(base: str, timeout: float) -> list[dict[str, Any]]:
    """Validate real usage statistics and the sidecar request logger."""
    started = time.perf_counter()
    allowed_users = {"lijuanyan", "buzhang", "caiwu", "caiwuzhuguan"}
    allowed_depts = {"采购业务部", "财务部"}
    try:
        login_status, login = _finance_request(
            base, "/api/login", {"username": "buzhang", "password": "88888888"}, timeout=timeout
        )
        token = str(login.get("token") or "")
        before_status, before = _finance_request(base, "/api/usage/stats?days=30", token=token, timeout=timeout)
        denied_status, _ = _finance_request(base, "/api/usage/stats?days=30", timeout=timeout)
        empty_status, empty = _finance_request(base, "/api/usage/stats?days=1", token=token, timeout=timeout)
        dashboard_status, _ = _finance_request(base, "/api/dashboard?project_id=demo_a", token=token, timeout=timeout)
        after_status, after = _finance_request(base, "/api/usage/stats?days=30", token=token, timeout=timeout)
        for _ in range(10):
            if int(after.get("total_actions") or 0) >= int(before.get("total_actions") or 0) + 1:
                break
            time.sleep(0.05)
            after_status, after = _finance_request(base, "/api/usage/stats?days=30", token=token, timeout=timeout)
        accounts = before.get("accounts") if isinstance(before.get("accounts"), list) else []
        departments = before.get("dept_dist") if isinstance(before.get("dept_dist"), list) else []
        actions = int(before.get("total_actions") or 0)
        active = int(before.get("active_users") or 0)
        average = before.get("avg_actions_per_user")
        numeric_avg = isinstance(average, (int, float)) and math.isfinite(float(average))
        math_ok = active <= actions and numeric_avg and (
            abs(float(average) - actions / active) < 0.05 if active else float(average) == 0.0
        )
        usernames = [str(row.get("username") or "") for row in accounts if isinstance(row, dict)]
        deps = {str(row.get("dept") or "") for row in departments if isinstance(row, dict)}
        elapsed = round((time.perf_counter() - started) * 1000, 1)
        checks = [
            {"case": "usage_stats_shape", "category": "usage", "question": "使用统计接口结构完整", "pass": before_status == 200 and all(key in before for key in ("total_actions", "active_users", "dept_count", "avg_actions_per_user", "history_unnamed_runs")), "failed_layers": [] if before_status == 200 and all(key in before for key in ("total_actions", "active_users", "dept_count", "avg_actions_per_user", "history_unnamed_runs")) else ["shape"], "elapsed_ms": elapsed},
            {"case": "usage_stats_no_fabrication", "category": "usage", "question": "账号列表只包含已配置账号", "pass": 0 < len(accounts) <= len(allowed_users) and set(usernames) <= allowed_users, "failed_layers": [] if 0 < len(accounts) <= len(allowed_users) and set(usernames) <= allowed_users else ["accounts"], "account_count": len(accounts), "elapsed_ms": elapsed},
            {"case": "usage_stats_math", "category": "usage", "question": "活跃用户和人均次数计算一致", "pass": math_ok, "failed_layers": [] if math_ok else ["math"], "elapsed_ms": elapsed},
            {"case": "usage_logging_works", "category": "usage", "question": "真实仪表盘请求会写入使用日志", "pass": dashboard_status == 200 and after_status == 200 and int(after.get("total_actions") or 0) >= actions + 1, "failed_layers": [] if dashboard_status == 200 and after_status == 200 and int(after.get("total_actions") or 0) >= actions + 1 else ["logging"], "elapsed_ms": elapsed},
            {"case": "usage_dept_from_real_roles", "category": "usage", "question": "部门分布来自真实角色映射", "pass": deps <= allowed_depts, "failed_layers": [] if deps <= allowed_depts else ["department"], "departments": sorted(deps), "elapsed_ms": elapsed},
            {"case": "usage_empty_window_safe", "category": "usage", "question": "一天窗口返回有效统计数值", "pass": empty_status == 200 and isinstance(empty.get("avg_actions_per_user"), (int, float)) and math.isfinite(float(empty["avg_actions_per_user"])), "failed_layers": [] if empty_status == 200 and isinstance(empty.get("avg_actions_per_user"), (int, float)) and math.isfinite(float(empty["avg_actions_per_user"])) else ["empty_window"], "elapsed_ms": elapsed},
            {"case": "usage_stats_requires_manager", "category": "usage", "question": "使用统计只对经理角色开放", "pass": login_status == 200 and bool(token) and denied_status == 401, "failed_layers": [] if login_status == 200 and token and denied_status == 401 else ["auth_required"], "elapsed_ms": elapsed},
        ]
        return checks
    except (OSError, ValueError, json.JSONDecodeError, KeyError, TypeError) as exc:
        return [{"case": "usage_suite", "category": "usage", "question": "应用使用情况回归", "pass": False, "failed_layers": ["request"], "error": str(exc), "elapsed_ms": round((time.perf_counter() - started) * 1000, 1)}]


def _evaluate_fee_waterfall(base: str, timeout: float) -> list[dict[str, Any]]:
    """Validate the read-only structured repayment waterfall against engine output."""
    started = time.perf_counter()
    try:
        payload = _fee_request(base, timeout)
        logs = payload.get("pay_log") or []
        payments = payload.get("pay_steps") or []
        present = bool(payments) and len(payments) == len(logs)
        conservation = True
        sourced = True
        for payment in payments:
            amount = float(payment.get("amount") or 0)
            steps = payment.get("steps") or []
            applied = sum(float(step.get("amount") or 0) for step in steps
                          if step.get("kind") in {"fee", "principal", "balance"})
            conservation = conservation and abs(applied - amount) <= 0.01
            sourced = sourced and all((step.get("ref") or {}).get("source")
                                      for step in steps if step.get("kind") in {"fee", "principal", "balance"})
        elapsed = round((time.perf_counter() - started) * 1000, 1)
        return [
            {"case": "fee_waterfall_present", "category": "fee-waterfall", "question": "每笔回款均有结构化冲抵流水", "pass": present, "failed_layers": [] if present else ["pay_steps"], "elapsed_ms": elapsed},
            {"case": "fee_waterfall_conservation", "category": "fee-waterfall", "question": "冲抵步骤金额守恒", "pass": present and conservation, "failed_layers": [] if present and conservation else ["conservation"], "elapsed_ms": elapsed, "grand_total": payload.get("grand_total")},
            {"case": "fee_waterfall_sources", "category": "fee-waterfall", "question": "冲抵步骤可溯源", "pass": present and sourced, "failed_layers": [] if present and sourced else ["source"], "elapsed_ms": elapsed},
        ]
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        return [{"case": "fee_waterfall_suite", "category": "fee-waterfall", "question": "资金占用费冲抵流水回归", "pass": False, "failed_layers": ["request"], "error": str(exc), "elapsed_ms": round((time.perf_counter() - started) * 1000, 1)}]


def _evaluate_audit(base: str, timeout: float) -> list[dict[str, Any]]:
    started = time.perf_counter()
    try:
        first = _audit_request(base, timeout, force=True)
        # 第二次读取缓存，验证历史对比且避免在离线/模型超时环境中重复跑完整 OCR。
        second = _audit_request(base, timeout, force=False)
        first_verdicts = first.get("agent_verdicts") or []
        second_verdicts = second.get("agent_verdicts") or []
        first_nodes = first.get("audit_trace", {}).get("nodes") or []
        second_nodes = second.get("audit_trace", {}).get("nodes") or []
        first_tools = [str(n.get("detail") or "") for n in first_nodes if "→" in str(n.get("detail") or "")]
        tools_by_issue: dict[Any, list[str]] = {}
        for node in first_nodes:
            detail = str(node.get("detail") or "")
            if "→" in detail:
                tools_by_issue.setdefault(node.get("issue_index", 0), []).append(detail)
        duplicate_tools = any(len(items) != len(set(items)) for items in tools_by_issue.values())
        verdict_order_same = [v.get("file") for v in first_verdicts] == [v.get("file") for v in second_verdicts]
        node_order_same = [n.get("detail") for n in first_nodes] == [n.get("detail") for n in second_nodes]
        history_present = bool(second.get("history")) and any(h.get("previous") for h in second.get("history", []))
        return [
            {"case": "audit_run_persist", "category": "audit", "question": "资料核查落库", "pass": bool(first.get("ok") and first.get("run_id")), "failed_layers": [] if first.get("ok") and first.get("run_id") else ["run_id"], "elapsed_ms": round((time.perf_counter() - started) * 1000, 1)},
            {"case": "audit_parallel_order", "category": "audit", "question": "资料核查并行保序", "pass": verdict_order_same and node_order_same, "failed_layers": [] if verdict_order_same and node_order_same else ["order"], "elapsed_ms": round((time.perf_counter() - started) * 1000, 1)},
            {"case": "audit_tool_dedup", "category": "audit", "question": "资料核查取证去重", "pass": not duplicate_tools, "failed_layers": [] if not duplicate_tools else ["duplicate_tool"], "tool_trace": first_tools, "elapsed_ms": round((time.perf_counter() - started) * 1000, 1)},
            {"case": "audit_history_compare", "category": "audit", "question": "资料核查历史对比", "pass": history_present, "failed_layers": [] if history_present else ["history"], "elapsed_ms": round((time.perf_counter() - started) * 1000, 1)},
        ]
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        return [{"case": "audit_suite", "category": "audit", "question": "资料核查回归", "pass": False, "failed_layers": ["request"], "error": str(exc), "elapsed_ms": round((time.perf_counter() - started) * 1000, 1)}]


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
    if args.category == "audit":
        audit_results = _evaluate_audit(args.base, args.timeout)
        if args.as_json:
            print(json.dumps({"base": args.base, "category": "audit", "total": len(audit_results), "passed": sum(bool(x.get("pass")) for x in audit_results), "results": audit_results}, ensure_ascii=False, indent=2))
        else:
            for result in audit_results:
                print(f"{'PASS' if result['pass'] else 'FAIL'} {result['case']} [{result.get('elapsed_ms', 0):.0f}ms] failed={','.join(result.get('failed_layers', []))}")
            print(f"资料核查回归: {sum(bool(x.get('pass')) for x in audit_results)}/{len(audit_results)}")
        return 0 if all(bool(x.get("pass")) for x in audit_results) else 1
    if args.category == "finance":
        finance_results = _evaluate_finance_intake(args.base, args.timeout)
        if args.as_json:
            print(json.dumps({"base": args.base, "category": "finance", "total": len(finance_results), "passed": sum(bool(x.get("pass")) for x in finance_results), "results": finance_results}, ensure_ascii=False, indent=2))
        else:
            for result in finance_results:
                print(f"{'PASS' if result['pass'] else 'FAIL'} {result['case']} [{result.get('elapsed_ms', 0):.0f}ms] failed={','.join(result.get('failed_layers', []))}")
            print(f"财务回款接口回归: {sum(bool(x.get('pass')) for x in finance_results)}/{len(finance_results)}")
        return 0 if all(bool(x.get("pass")) for x in finance_results) else 1
    if args.category == "perspective":
        perspective_results = _evaluate_perspective(args.base, args.timeout)
        if args.as_json:
            print(json.dumps({"base": args.base, "category": "perspective", "total": len(perspective_results), "passed": sum(bool(x.get("pass")) for x in perspective_results), "results": perspective_results}, ensure_ascii=False, indent=2))
        else:
            for result in perspective_results:
                print(f"{'PASS' if result['pass'] else 'FAIL'} {result['case']} [{result.get('elapsed_ms', 0):.0f}ms] failed={','.join(result.get('failed_layers', []))}")
            print(f"视角切换与权限回归: {sum(bool(x.get('pass')) for x in perspective_results)}/{len(perspective_results)}")
        return 0 if all(bool(x.get("pass")) for x in perspective_results) else 1
    if args.category == "usage":
        usage_results = _evaluate_usage(args.base, args.timeout)
        if args.as_json:
            print(json.dumps({"base": args.base, "category": "usage", "total": len(usage_results), "passed": sum(bool(x.get("pass")) for x in usage_results), "results": usage_results}, ensure_ascii=False, indent=2))
        else:
            for result in usage_results:
                print(f"{'PASS' if result['pass'] else 'FAIL'} {result['case']} [{result.get('elapsed_ms', 0):.0f}ms] failed={','.join(result.get('failed_layers', []))}")
            print(f"使用情况埋点回归: {sum(bool(x.get('pass')) for x in usage_results)}/{len(usage_results)}")
        return 0 if all(bool(x.get("pass")) for x in usage_results) else 1
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
    waterfall_results = _evaluate_fee_waterfall(args.base, args.timeout)
    results.extend(waterfall_results)
    if not args.as_json:
        for result in waterfall_results:
            print(f"{'PASS' if result['pass'] else 'FAIL'} {result['case']} "
                  f"[{result.get('elapsed_ms', 0):.0f}ms] failed={','.join(result.get('failed_layers', []))}")
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

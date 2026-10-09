# 绿金街供应链项目管理平台

> **在线体验版（免安装）**：打开网页即可自己点、自己看 ——
> 部署到 GitHub Pages 后，入口是 `https://<用户名>.github.io/lvjinjie-supply-chain-platform/`。
> 纯静态、不含后端与密钥；数据为预置快照，写操作为本地模拟。
> 想跑完整功能（实时检索 / DeepSeek 问答 / OCR），请看下面的部署说明。

---

# A项目演示供应链控制塔

一个面向“资金方介入型供应链中间商”的本地控制塔原型。

## 能做什么

- 统一看项目应收、未收、逾期、累计回款、资金占用费
- 按上月 21 日至本月 20 日自动归集结算周期
- 按合同规则计算次月月末、7 天宽限期和逾期日
- 录入 OA 模拟回款并自动冲抵余额
- 识别上游/下游送货单、对账单、发票、台账和资金占用费表
- 用 Ollama 做合同 RAG 问答
- 以 Router / Retriever / Auditor / Rule Engine / Allocator / Writer 的子 Agent 图组织执行

## 启动

```powershell
cd C:\Users\<user>\Documents\Codex\2026-09-24\wo
pip install -r requirements.txt   # 首次运行安装依赖
python server.py
```

使用交付包启动脚本时会自动检查并安装依赖；首次下载 OCR 组件可能需要 1–3 分钟，期间窗口没有响应并不代表服务崩溃。

打开 `http://127.0.0.1:8787`。

**资金占用费与回款**：扫描「回款回单」目录，对回单图片做本地 OCR 提取记账日期与金额，再按合同规则（本金×0.0002/天、回款日不计息、按「前期费用→旧本金→本期费用」顺序冲抵）自动核算月度确认单。OCR 模型随包内置，运行时无需联网。

## 接口

- `GET /api/dashboard`
- `GET /api/cycles`
- `GET /api/documents`
- `GET /api/agents`
- `GET /api/health`
- `POST /api/ask`
- `POST /api/repayments`
- `POST /api/refresh`
- `GET /api/reindex`

## 说明

当前版本没有直接依赖 `langchain` / `langgraph` 包，而是先做了一个可运行的本地图式实现，节点和返回结构已经按后续迁移到 LangChain / LangGraph 的方式设计好。Ollama 默认读取 `http://127.0.0.1:11434`，意图分类模型默认 `qwen2.5:1.5b`，向量模型 `bge-m3`，全部本地运行。

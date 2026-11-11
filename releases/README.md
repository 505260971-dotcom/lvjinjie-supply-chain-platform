# releases/ — 离线演示交付包

本目录**不再存放二进制压缩包**。演示交付包体积约 20 MB，若随仓库提交，会让每次 `git clone` 都要多拉 20 MB，而它本身是「源代码 + 演示资料」的打包快照，属于可再生成产物。

## 下载

| 版本 | 文件名 | 大小 | 下载 |
| --- | --- | --- | --- |
| v1.0.0-demo | `lvjinjie-supply-chain-platform-v1.0.0-demo.zip` | 19.8 MB | [GitHub Releases](https://github.com/505260971-dotcom/lvjinjie-supply-chain-platform/releases/tag/v1.0.0-demo) |

校验值（MD5）：`47f69cf35dad7330e6d13aeeb2781ec6`

## 包内含什么

- 完整前后端（`server.py` + `index.html`）与预置演示数据（10 个项目，A 项目为主演示链路）
- 图文说明书、路演速查卡、评委请先读我
- 评测脚本 `agent_harness.py`（35 条回归基线）

## 怎么用

解压后进入目录执行 `python server.py`，浏览器打开 <http://127.0.0.1:8787>，演示账号 `lijuanyan` / `88888888`。

无需任何模型 Key：不配置 `DEEPSEEK_API_KEY` 且本机没有 ollama 时，系统自动走规则引擎路径，评测基线依然 35/35。

## 为什么不入库

- 内容是仓库当前源码 + 演示资料的重打包，纳入版本库会与源码重复；
- 每次改代码都要同步重新打包，二进制文件在 git 里会持续膨胀（哪怕删掉，历史上仍占体积）；
- GitHub Releases 对单个资产支持到 2 GB，且提供稳定的下载链接与版本归档。

需要重新打包时，按「源代码目录 + 演示资料目录 + 启动脚本」结构压缩为 zip，用新 tag 发布到 Releases，并在此表补一行。

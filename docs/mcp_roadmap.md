# MCP Integration Roadmap

## 📖 概述

本文檔規劃 LLM Office I/O 與 **Model Context Protocol (MCP)** 的整合計畫。

**MCP** 是由 Anthropic 推出的標準化協議，用於連接 LLM 與外部工具和數據源。

---

## 🎯 目標

### 短期目標（已完成 ✅）
- [x] 創建 LLM-friendly API 層
- [x] 統一返回格式
- [x] JSON 輸入/輸出支援
- [x] 工具描述文檔（JSON Schema）

### 長期目標（v2.0.0）
- [ ] 完整 MCP Server 實作
- [ ] 支援 MCP 協議
- [ ] 工具自動發現
- [ ] Streaming 支援
- [ ] 並發請求處理

---

## 🏗️ 架構設計

### 當前架構（v1.3.0）

```
┌─────────────┐
│  AI Agent   │
└─────┬───────┘
      │ Python API
      ▼
┌─────────────┐
│  llm_api.py │  ← 簡化API層
└─────┬───────┘
      │
      ▼
┌────────────────────────────────┐
│  word_editor | ppt_editor      │
│  excel_editor | batch_processor│
└────────────────────────────────┘
```

### 目標架構（v2.0.0）

```
┌─────────────┐
│  AI Agent   │
└─────┬───────┘
      │ MCP Protocol (JSON-RPC)
      ▼
┌──────────────────┐
│   MCP Server     │  ← 標準化接口
│  (office-tools)  │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│   llm_api.py     │  ← 現有簡化層
└────────┬─────────┘
         │
         ▼
┌──────────────────────────────┐
│  Core Editors (現有實作)     │
└──────────────────────────────┘
```

---

## 📋 實作計畫

### Phase 1: MCP Server 骨架（4小時）

**檔案**: `src/mcp_server.py`

```python
from mcp.server import Server
from mcp.types import Tool, Resource, Prompt
import asyncio

app = Server("office-tools")

@app.list_tools()
async def list_tools() -> list[Tool]:
    """列出所有可用工具"""
    return [
        Tool(
            name="office_replace_text",
            description="替換Office文檔中的文字",
            inputSchema={...}
        ),
        # 更多工具...
    ]

@app.call_tool()
async def call_tool(name: str, arguments: dict):
    """執行工具調用"""
    from .llm_api import execute_command
    result = execute_command(...)
    return result
```

---

### Phase 2: MCP 協議實作（6小時）

**功能**:
1. **JSON-RPC 2.0** 支援
2. **Transport 層**: stdio, HTTP, WebSocket
3. **請求/響應** 處理
4. **錯誤標準化**: MCP 錯誤碼

**範例請求**:
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "office_replace_text",
    "arguments": {
      "file_path": "report.docx",
      "old_text": "2024",
      "new_text": "2025"
    }
  }
}
```

**範例響應**:
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "result": {
    "content": [
      {
        "type": "text",
        "text": "成功替換 5 處"
      }
    ],
    "isError": false
  }
}
```

---

### Phase 3: 資源支援（4小時）

MCP 支援 **Resources**，讓 LLM 可以讀取文檔內容：

```python
@app.list_resources()
async def list_resources():
    """列出可用資源"""
    return [
        Resource(
            uri="office://documents",
            name="Office Documents",
            description="List available Office files",
            mimeType="application/json"
        )
    ]

@app.read_resource()
async def read_resource(uri: str):
    """讀取資源內容"""
    if uri == "office://documents":
        # 列出可用文檔
        files = glob.glob("*.docx") + glob.glob("*.pptx")
        return {"files": files}
    
    # 讀取特定文檔
    if uri.startswith("office://doc/"):
        filepath = uri.replace("office://doc/", "")
        # 提取文檔內容...
```

---

### Phase 4: Prompts 支援（2小時）

提供預定義的操作範本：

```python
@app.list_prompts()
async def list_prompts():
    return [
        Prompt(
            name="batch_update_year",
            description="批次更新所有報告的年份",
            arguments=[
                {"name": "pattern", "description": "檔案模式"},
                {"name": "old_year", "description": "舊年份"},
                {"name": "new_year", "description": "新年份"}
            ]
        )
    ]

@app.get_prompt()
async def get_prompt(name: str, arguments: dict):
    if name == "batch_update_year":
        # 生成操作序列
        return {
            "messages": [
                {
                    "role": "user",
                    "content": f"批次將 {arguments['pattern']} 中的 {arguments['old_year']} 替換為 {arguments['new_year']}"
                }
            ]
        }
```

---

### Phase 5: 進階功能（8小時）

#### 5.1 Streaming 支援
```python
@app.call_tool_streaming()
async def call_tool_streaming(name: str, arguments: dict):
    """支援串流式回應"""
    for progress in process_files():
        yield {
            "type": "progress",
            "data": progress
        }
```

#### 5.2 並發請求
```python
import asyncio

async def handle_concurrent_requests():
    """同時處理多個請求"""
    tasks = [
        call_tool("replace_text", {...}),
        call_tool("add_image", {...})
    ]
    results = await asyncio.gather(*tasks)
    return results
```

#### 5.3 Session 管理
```python
class SessionManager:
    """管理多個客戶端連接"""
    def __init__(self):
        self.sessions = {}
    
    async def create_session(self, client_id: str):
        self.sessions[client_id] = {
            "open_documents": {},
            "history": []
        }
```

---

## 📦 依賴套件

```txt
# requirements-mcp.txt
mcp>=1.0.0           # MCP Python SDK
pydantic>=2.0.0      # 數據驗證
asyncio              # 異步支援
websockets>=12.0     # WebSocket 支援
aiohttp>=3.9.0       # HTTP 支援
```

---

## 🧪 測試策略

### 單元測試
```python
# tests/test_mcp_server.py
async def test_list_tools():
    server = create_test_server()
    tools = await server.list_tools()
    assert len(tools) > 0
    assert tools[0].name == "office_replace_text"

async def test_call_tool():
    server = create_test_server()
    result = await server.call_tool(
        "office_replace_text",
        {"file_path": "test.docx", ...}
    )
    assert result["success"] == True
```

### 整合測試
```python
async def test_mcp_client_integration():
    # 使用 MCP 客戶端測試完整流程
    from mcp.client import Client
    
    async with Client("office-tools") as client:
        tools = await client.list_tools()
        result = await client.call_tool("office_replace_text", {...})
```

---

## 📊 階段時程

| 階段 | 任務 | 時間 | 優先級 |
|------|------|------|--------|
| 1 | MCP Server 骨架 | 4h | P0 |
| 2 | 協議實作 | 6h | P0 |
| 3 | 資源支援 | 4h | P1 |
| 4 | Prompts 支援 | 2h | P2 |
| 5 | 進階功能 | 8h | P2 |
| **總計** | | **24h** | |

---

## 🎯 里程碑

### v1.4.0 - MCP Alpha（預計 1 個月）
- [x] 基礎 LLM API（已完成）
- [ ] MCP Server 骨架
- [ ] 基本工具調用

### v1.5.0 - MCP Beta（預計 2 個月）
- [ ] 完整協議支援
- [ ] 資源讀取
- [ ] Prompts 支援

### v2.0.0 - MCP GA（預計 3 個月）
- [ ] Streaming
- [ ] 並發處理
- [ ] 生產級穩定性

---

## 💡 使用範例（未來）

### Claude Desktop 整合

```json
// claude_desktop_config.json
{
  "mcpServers": {
    "office-tools": {
      "command": "python",
      "args": ["-m", "src.mcp_server"],
      "env": {}
    }
  }
}
```

### 直接使用

```python
# 啟動 MCP Server
python -m src.mcp_server --transport stdio

# 或
python -m src.mcp_server --transport http --port 8080
```

---

## 📚 參考資源

- [MCP 官方文檔](https://modelcontextprotocol.io/)
- [MCP Python SDK](https://github.com/anthropics/mcp-python)
- [MCP 規範](https://spec.modelcontextprotocol.io/)

---

## ⚠️ 注意事項

1. **向後兼容**: MCP 層建立在現有 API 之上，不影響現有功能
2. **漸進式遷移**: 可以先支援部分工具，逐步擴展
3. **效能考量**: 需要測試並發效能和記憶體使用
4. **安全性**: 需要添加認證和授權機制

---

**文檔版本**: 1.0  
**最後更新**: 2025-12-02  
**負責人**: Development Team

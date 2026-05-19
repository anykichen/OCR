# OcrTcpToKeyboard v2

> 从原版完全重写，去掉 Interception 驱动依赖，完整支持 Excel / Word / PowerPoint / WPS / 任意 Unicode / 中文 注入。

---

## 原版问题分析

| 问题 | 原因 |
|------|------|
| Excel/Word/PPT 注入全部失败 | EM_REPLACESEL 只对 Win32 Edit 控件有效；Office 是 COM/自绘控件 |
| 中文/Unicode 注入失败 | ScanCode 方式只支持 ASCII，无法模拟非英文字符 |
| 需要安装 interception.dll 驱动 | 内核级键盘驱动，需要管理员 + 特殊安装步骤，维护复杂 |

---

## v2 解决方案：双策略注入

```
接收文字 (TCP)
     │
     ▼
检测当前焦点窗口 class
     │
     ├──► Win32 Edit 控件 (Notepad, 标准文本框)
     │         └──► EM_REPLACESEL 直接注入（快速，不影响剪贴板）
     │
     └──► 其他所有窗口 (Excel, Word, PPT, WPS, 浏览器, 任意...)
               └──► 剪贴板注入 (Clipboard + Ctrl+V + 可选 Enter)
                         ├── 保存当前剪贴板内容
                         ├── 写入目标文字
                         ├── 发送 Ctrl+V
                         ├── 可选：发送 Enter
                         └── 异步还原剪贴板（600ms 后）
```

### 为什么剪贴板方式对 Office 最稳定？

- Excel/Word/PPT 内部是 COM + 自绘，没有标准 Win32 消息队列可以轻易注入
- Clipboard + Ctrl+V 是 Windows 标准粘贴流程，和用户手动 Ctrl+V 完全等价
- 完整支持：中文、日文、韩文、emoji、任意 Unicode 字符
- 无需驱动、无需管理员权限（SendKeys 除外）

---

## 编译方法

### 方法一：Visual Studio（推荐）

1. 用 VS 2022 打开 `OcrTcpToKeyboard2.csproj`
2. 目标框架选 `.NET Framework 4.8`（或改为 `net6.0-windows` 也可）
3. `Ctrl+B` 编译，`bin\Release\` 下得到 `.exe`

### 方法二：命令行

```bat
# 需要安装 .NET SDK 6+
dotnet build -c Release
```

### 方法三：改用 .NET 6+（可选）

将 `.csproj` 中的 `<TargetFramework>net48</TargetFramework>` 改为：
```xml
<TargetFramework>net8.0-windows</TargetFramework>
```
重新编译即可，代码无需修改。

---

## TCP 协议

- 监听地址：`127.0.0.1:62020`（与原版兼容）
- 发送：UTF-8 文本（支持中文 / 任意 Unicode）
- 回复：`OK\n` 或 `FAIL\n`

```python
# Python 测试示例
import socket
s = socket.socket()
s.connect(('127.0.0.1', 62020))
s.send('你好世界，这是测试文字'.encode('utf-8'))
print(s.recv(64))  # b'OK\n'
s.close()
```

---

## 使用说明

1. 运行 `OcrTcpToKeyboard2.exe`
2. 点击 **启动监听**
3. 在 Excel / Word / PPT 等软件中，**点击一下目标单元格/文本框**（确保该窗口处于前台且有光标）
4. 外部程序通过 TCP 发送文字 → 自动粘贴到当前光标位置

### "末尾追加回车" 选项

- ✅ 勾选：每次注入后自动按一次 Enter（适合 Excel 逐行填写、聊天框发送等）
- ☐ 不勾选：纯文字注入，不自动换行（适合 Word / PPT 连续输入）

---

## 注意事项

- Excel 注入前请确保目标单元格处于**编辑模式**（双击或 F2），否则 Ctrl+V 会替换整格
- 也可以不进入编辑模式，直接在选中单元格时粘贴（Excel 会直接写入单元格值）
- 注入时会**短暂占用剪贴板**（约 600ms 后恢复原内容），极端情况下可能干扰同时操作剪贴板的其他操作
- 若目标软件有"安全粘贴确认"（如微信 PC 版），可能需要手动确认

---

## 与原版对比

| 功能 | 原版 | v2 |
|------|------|----|
| Notepad 注入 | ✅ | ✅ |
| Excel 注入 | ❌ | ✅ |
| Word 注入 | ❌ | ✅ |
| PPT 注入 | ❌ | ✅ |
| WPS 注入 | ❌ | ✅ |
| 中文/Unicode | ❌ | ✅ |
| 需要安装驱动 | ✅（必须） | ❌（无需） |
| 需要管理员权限 | ✅ | ❌ |
| 剪贴板保护 | — | ✅（异步还原） |

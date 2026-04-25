# Windows Desktop Control Skill

## 概述
本 Skill 通过调用 `control.ps1` 脚本实现高度拟人化的鼠标操作、键盘输入和屏幕截图。  
**核心能力**：结合 OpenClaw 的图片理解能力，完成“截图→视觉定位→操作”的自动化闭环，无需额外视觉模型。

## 鼠标轨迹特性（Agent 须知）
- 移动命令 (`move`, `click` 等) 内部采用了多段随机缓动、布朗游走震颤、动态步数与逐帧延迟抖动。
- 长距离 (>300px) 自动拆分为快速扫视 + 停顿 + 精细校正，模拟视觉追踪滞后。
- 这些特性使轨迹在统计上极难与真人区分，Agent 在执行高频点击、拖拽、签到等任务时可以放心使用。

## 前置条件
- 操作系统：Windows 10/11，PowerShell 5.1 及以上。
- **`control.ps1` 已加入系统 PATH**（可在任意 PowerShell 会话中直接运行 `control.ps1`）。若未配置，请先阅读《环境配置指南》完成 PATH 设置。
- 工作目录：截图固定保存在 `$HOME\.openclaw\Workspace\last_screen.png`，无需手动创建。

## 通用命令格式
所有命令均通过 PowerShell 执行，脚本名直接使用 `control.ps1`（因已配环境变量）：
```
powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File control.ps1 -action <动作> [其他参数]
```

> **重要提示**：下文所有“示例”中的坐标、文本内容均为演示占位，**严禁直接复制使用**，必须根据实际识别结果或用户意图替换为有效值。

---

## 1. 截图

**命令**：
```
powershell -File control.ps1 -action snapshot
```
**返回**：截图的绝对路径（字符串）。  
*示例*（返回的路径会根据系统实际路径变化，无需固定）：
```
C:\Users\YourName\.openclaw\Workspace\last_screen.png
```
- 若返回以 `error:` 开头，表示截图失败，请检查权限或磁盘空间。

---

## 2. 鼠标移动

**命令**：
```
powershell -File control.ps1 -action move -x <整数> -y <整数>
```
- `<整数>` 为目标屏幕像素坐标。
- 行为：拟人化轨迹移动到目标点，超出屏幕范围会自动钳位。

*示例（请勿直接使用，坐标需根据实际情况替换）*：
```
powershell -File control.ps1 -action move -x 500 -y 300
```

---

## 3. 鼠标点击

所有点击操作均会先移动鼠标到目标坐标，再执行点击，带有随机延迟。

| 动作 | 所需参数 | 命令结构 |
|------|----------|----------|
| 左键单击 | `-action click -x <x> -y <y>` | `powershell ... -action click -x <x> -y <y>` |
| 右键单击 | `-action rightclick -x <x> -y <y>` | 同上 |
| 左键双击 | `-action doubleclick -x <x> -y <y>` | 同上 |
| 中键单击 | `-action middleclick -x <x> -y <y>` | 同上 |

*示例（坐标仅为示意，实际应替换为视觉定位结果）*：
```
powershell -File control.ps1 -action click -x 640 -y 480        # 左键单击示例
powershell -File control.ps1 -action rightclick -x 640 -y 480    # 右键单击示例
powershell -File control.ps1 -action doubleclick -x 640 -y 480   # 双击示例
powershell -File control.ps1 -action middleclick -x 640 -y 480   # 中键示例
```

---

## 4. 鼠标滚轮

**命令**：
```
powershell -File control.ps1 -action wheel -scroll <整数>
```
- `<整数>` 正值上滚，负值下滚；通常一个刻度为 `120` 或 `-120`。

*示例（请根据实际滚动方向和格数调整数值）*：
```
powershell -File control.ps1 -action wheel -scroll 360      # 上滚 3 格示例
powershell -File control.ps1 -action wheel -scroll -240     # 下滚 2 格示例
```

---

## 5. 键盘操作

### 5.1 系统快捷键（无需坐标）

| 功能 | 命令（直接复制可用） |
|------|---------------------|
| 全选 (Ctrl+A) | `powershell -File control.ps1 -action selectall` |
| 复制 (Ctrl+C) | `powershell -File control.ps1 -action copy` |
| 粘贴 (Ctrl+V) | `powershell -File control.ps1 -action paste` |

**注意**：以上三个命令无额外参数，可直接使用，无示例占位。

---

### 5.2 模拟真人打字

**命令**：
```
powershell -File control.ps1 -action type -text "<文本内容>"
```

**规则**：
- 文本必须用双引号包裹；内容中包含双引号时用反引号转义：`` `" ``。
- 支持所有字母（自动大小写）、数字、常见符号及控制字符（详见附录）。
- 换行建议使用 `` `r`n `` 模拟 Windows 回车换行。

*示例（文本内容均为演示，请勿直接复制，应根据用户实际要输入的字符串动态生成）*：
```
powershell -File control.ps1 -action type -text "Hello World"                    # 普通文本示例
powershell -File control.ps1 -action type -text "密码：`t123456"                 # 包含 Tab 示例
powershell -File control.ps1 -action type -text "错误`b`b修正"                   # 退格修正示例
powershell -File control.ps1 -action type -text "他说：`"你好`""                 # 含双引号示例
powershell -File control.ps1 -action type -text "第一行`r`n第二行"               # 换行示例
```

---

## 6. 屏幕分辨率与坐标转换

- 所有 `-x`、`-y` 必须为**绝对像素坐标**，原点为主显示器左上角。
- 若 OpenClaw 从图片分析得到的是归一化坐标（0~1），需先获取屏幕分辨率：

**获取分辨率命令（直接可用）**：
```
powershell -Command "Add-Type -AssemblyName System.Windows.Forms; $b=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds; \"$($b.Width) $($b.Height)\""
```
返回示例：`1920 1080` （实际值因设备而异）。

**换算方法**（需在代码或思维中执行，非命令）：
```
abs_x = round(norm_x * width)
abs_y = round(norm_y * height)
```
并将结果限定在 `[0, 分辨率-1]` 范围内。

---

## 7. 视觉定位自动化流程（核心）

当用户要求“点击屏幕上的某个元素（如图标、按钮、输入框）”时，请按以下步骤**依次调用**：

1. **截图**  
   执行 `powershell -File control.ps1 -action snapshot`，获取最新截图文件路径。

2. **读图与定位**  
   使用 OpenClaw 的图片理解能力读取该截图，根据目标描述（例如“Chrome 图标”）找到目标中心点的**屏幕绝对坐标 (x, y)**。  
   - 如果直接输出绝对像素坐标，直接使用。
   - 如果输出归一化坐标，按第 6 节转换为绝对像素。

3. **执行鼠标操作**  
   将获得的坐标填入对应的点击命令，例如单击或双击。

4. **结果验证**  
   可再次截图检查界面是否变化，若点击未生效，允许以 2~5 像素微调坐标后重试一次。

> **请勿使用文档中的示例坐标 (640,480) 或 (320,540) 执行真实点击**，这些仅为说明格式的占位数字，真实值必须来自视觉分析结果。

---

## 8. 错误处理与注意事项

- **脚本未找到**：若返回 `error: 未找到 control.ps1`，检查 PATH 是否生效，必要时重启终端或 OpenClaw 环境。
- **权限不足**：一般无需管理员，若失败可尝试以管理员身份运行 OpenClaw。
- **并发冲突**：脚本内部已包含随机延时（模拟真人），请顺序执行命令，避免快速连续调用。
- **组合键安全**：全部快捷键采用 `try/finally` 确保释放，断电或中断不会卡键。
- **调试模式**：任何命令追加 `-Verbose` 可输出详细日志，例如：
  ```
  powershell -File control.ps1 -action click -x 500 -y 300 -Verbose   # 仅为 Verbose 用法示例，坐标需替换
  ```

---

## 9. 支持的字符列表（完整）

**字母**：a-z（大小写自动处理）  
**数字**：0-9  
**符号**：`` ` ! @ # $ % ^ & * ( ) - _ = + [ ] { } \ | ; : ' " , . < > / ? ~ ``  
**控制字符**：空格、Tab（`` `t ``）、回车（`` `r ``）、换行（`` `n ``）、退格（`` `b ``）

不在上述列表中的字符会被静默忽略。

---

## 10. 场景模板

### 打开桌面程序
1. 截图找到图标 → 2. 双击坐标（`-action doubleclick -x <定位X> -y <定位Y>`）

### 在输入框中输入文本
1. 截图找到输入框 → 2. 单击激活输入框（`-action click -x <X> -y <Y>`） → 3. 用 `type` 输入内容

### 复制选中文文本（无原生拖拽，可用的替代方案）
1. 全选：`-action selectall`
2. 复制：`-action copy`

其他组合操作可参照上述命令自由搭配，始终以视觉定位结果为实际参数。

---

**版本**：1.1 · **日期**：2026-04-25  
**维护者**：Jakechenvb

# Windows Desktop Control Skill 环境配置指南

本 Skill 依赖 `control.ps1` 脚本实现桌面控制。使用前需将脚本所在目录添加至系统环境变量 `PATH`，以便 OpenClaw 在任意位置直接调用脚本。

## 1. 确认 control.ps1 存放位置

将 `control.ps1` 放在一个固定的文件夹中（例如 `C:\Users\你的用户名\.openclaw\Workspace`）。  
**记下该文件夹的完整路径**，后续步骤需要使用。

## 2. 将脚本目录加入 PATH（推荐添加至用户环境变量）

### 图形界面方式
1. 按 `Win + R`，输入 `sysdm.cpl` 并回车，打开“系统属性”。  
2. 切换到 **高级** 选项卡，点击 **环境变量(N)**。  
3. 在 **用户变量** 区域中找到 `Path` 变量，选中后点击 **编辑**。  
4. 点击 **新建**，粘贴你的脚本文件夹路径（例如 `C:\Users\你的用户名\.openclaw\Workspace`）。  
5. 依次点击 **确定** 保存所有窗口。

### PowerShell 命令行方式（推荐）
以当前用户权限打开 PowerShell，执行以下命令（请替换为你的实际路径）：
```powershell
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Users\你的用户名\.openclaw\Workspace", "User")
```
执行完成后，**关闭并重新打开**所有 PowerShell 窗口，使新 PATH 生效。

## 3. 验证环境变量

打开一个新的 PowerShell 窗口，直接输入：
```powershell
control.ps1
```
如果输出类似以下提示信息，表示 PATH 配置成功：
```
请指定 -action 参数: move/click/rightclick/doubleclick/middleclick/wheel/snapshot/type/selectall/copy/paste
```
如果提示“无法将 control.ps1 识别为 cmdlet...”，请检查路径是否正确，并确认已重启终端。

## 4. 配置 Skill 文件

1. 将仓库中提供的 `skill.md` 文件放入 OpenClaw 的 Skill 目录（通常为项目根目录下的 `skills/` 文件夹）。  
2. 该 `skill.md` 已默认使用 `control.ps1` 作为命令（无需路径），因为已经将脚本目录加入 `PATH`，所以**无需手动修改 skill.md 中的任何命令路径**。  
3. 如果 OpenClaw 运行意外报错“未找到 control.ps1”，请再次确认环境变量是否在当前会话中生效（可运行 `$env:Path` 检查是否包含对应目录）。

## 5. 快速测试

在 OpenClaw 中尝试以下指令，验证 Skill 是否正常工作：
- “截取当前屏幕”  
- “移动鼠标到 500,300”  
- “在桌面上找到回收站并打开”

如果能成功执行截图、移动和视觉点击，说明整个环境已就绪。

## 常见问题

**Q：为什么配置了 PATH 后还是提示找不到脚本？**  
A：请重启所有已打开的 PowerShell 或终端窗口，系统环境变量的变更需要新会话才能加载。亦可尝试重启 OpenClaw 服务。

**Q：我可以将脚本放在其他路径吗？**  
A：可以，只要将该路径加入 `PATH` 即可。如果不想修改 `PATH`，也可以在 `skill.md` 中将所有 `control.ps1` 替换为脚本的绝对路径（如 `"C:\...\control.ps1"`）。

**Q：截图保存在哪里？**  
A：截图固定保存在 `C:\Users\你的用户名\.openclaw\Workspace\last_screen.png`，无需手动创建目录。

## 附：相关文件说明

| 文件 | 用途 |
|------|------|
| `control.ps1` | 桌面控制核心脚本，包含鼠标、键盘、截图功能 |
| `skill.md` | OpenClaw Skill 描述文件，定义如何调用脚本实现自动化 |
| `README.md` | 本文件，环境安装与配置指南 |

# ForceOneExplorerWindow

Windows 11 资源管理器多窗口自动合并为单窗口多标签页的 AutoHotkey v2 脚本（[1explorer.ahk](1explorer.ahk)）。

基于 Reddit 上的[原始版本](https://www.reddit.com/r/AutoHotkey/comments/1fvsm2y/force_windows_11_to_open_file_explorer_in_new_tab/)，针对中文系统和控制面板窗口做了大量修正。

## 用法

1. 安装 [AutoHotkey v2](https://www.autohotkey.com/)。
2. 双击 `1explorer.ahk` 运行；如需开机自启，放入 `shell:startup` 文件夹。
3. 运行后：新打开的资源管理器窗口会自动合并为第一个窗口的标签页。

## 工作原理

通过 `SetWinEventHook` 监听 `CabinetWClass`（资源管理器窗口）的系统事件：

| 事件 | 处理 |
|---|---|
| `EVENT_OBJECT_CREATE` | 尽早把新窗口设为全透明（视觉隐藏），避免闪烁 |
| `EVENT_OBJECT_SHOW` | 第一个窗口作为主窗口 `FirstWindow`；其余窗口延迟进入 `MergeWindow` 合并流程 |
| `EVENT_OBJECT_DESTROY` | 清理忽略列表 |

`MergeWindow` 合并流程：

1. `GetPath` 通过 `Shell.Application.Windows` + `IShellBrowser` 取得窗口当前路径（带重试，等 COM 就绪）。
2. `IsExcludedPath` 判断该路径是否不可合并（见下节）。
3. `OpenInNewTab` 向主窗口发送 `WM_COMMAND 0xA21B` 新建标签页，再用 `IWebBrowser2.Navigate2` 导航到目标路径：
   - 普通文件路径 → 直接 `Navigate2(path)`；
   - 虚拟文件夹（`::{GUID}`，含 `::{父}\::{子}` 复合路径）→ `SHParseDisplayName` 解析为 PIDL，包装成 `VT_ARRAY|VT_UI1` 的 VARIANT 后 `Navigate2`。
4. **弹窗检测**：导航后等 1 秒对比窗口列表快照。某些页（如"管理存储空间"）无法原地导航，系统会另外弹出一个独立窗口——检测到新窗口即判定该页不可合并。
5. **宽限确认**：合并看似成功也不立即关闭原窗口，而是等 2 秒（`PendingMerges` + `FinalizeMerge`）。期间若有同路径新窗口迟到弹出，则关闭弹出的复制品、恢复**原来的 Explorer 窗口**（避免窗口"跑到控制面板"图标下）；确认无弹窗才关闭原窗口完成合并。

任何一步失败（取不到路径、解析失败、导航失败、弹出新窗口）都会恢复原窗口显示并记入忽略列表——**不报错、不丢窗口**。

## 控制面板页的处理策略（三层）

判断一律基于**路径**而非窗口文字——窗口文字在 `EVENT_OBJECT_SHOW` 时尚未填充，存在时序竞争；且中文系统显示文本不同。附带好处：文件夹名称里含"控制面板"也能正常合并。

1. **控制面板主页 / 所有控制面板项**：完整路径排除（`::{26EE0668-…}`、`::{21EC2020-…}`），保持独立窗口。
2. **已知"导航会弹新窗口"的页**：按**末尾 CLSID** 匹配直接跳过（入口不同，复合路径的父 GUID 不同，故不匹配完整路径）。目前名单：
   - `{F942C606-0914-47AB-BE56-1321B8035096}` 管理存储空间
3. **未知弹窗页**：运行时自动识别（第 4 步弹窗检测），记入 `BadPaths`，本次运行内再次打开时直接跳过。

### 如何添加新的已知弹窗页

发现某个控制面板页不能合并时，先保持该窗口打开，用 PowerShell 查询其路径：

```powershell
(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Document.Folder.Self.Path }
```

取输出的最后一段 CLSID，加到 `IsExcludedPath` 中 `knownBad` 的 `(?:...)` 列表里即可（本次运行已自动识别的页重启脚本后仍需加入名单才会直接跳过）。

控制面板项 CLSID 与名称对照可用以下命令查询：

```powershell
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace' | ForEach-Object { $guid = $_.PSChildName; $name = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\CLSID\$guid" -ErrorAction SilentlyContinue).'(default)'; "$guid  $name" }
```

## 已知限制

- 首次遇到某个未知弹窗页时，主窗口会多出一个空白"主页"标签页，需手动关闭；第二次打开该页起不再出现。
- `BadPaths`（运行时自动识别的名单）只在本次运行有效，重启脚本后需重新识别一次；常遇到的页建议按上节方法加入 `knownBad`。
- 合并成功后的 2 秒宽限期内，如果用户手动再次打开**同一路径**的窗口，可能被误判为"弹窗"而关闭并把该路径标记为不可合并（极少见，重启脚本恢复）。
- 若第一个打开的窗口本身就是控制面板主页，`CheckFirstWindow` 会将其降级为独立窗口，后续文件夹窗口另立主窗口。

## 关键实现

- 隐藏方式用 `WinSetTransparent(0)` 而非隐藏窗口，规避 `DetectHiddenWindows` 的干扰。
- `WaitForSameWindowCount` 等待 `Shell.Application.Windows` 数量与标签控件数量一致（忽略已排除窗口），保证 `GetPath`/`OpenInNewTab` 按索引取到正确的 shell 窗口。
- `MergeLock` 保证一次只合并一个窗口——弹窗检测依赖窗口列表快照，并发合并会互相干扰。
- 原窗口关闭前始终处于透明隐藏状态，因此 1 秒弹窗检测和 2 秒宽限期用户均无感知。

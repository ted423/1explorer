# ForceOneExplorerWindow

[中文](#中文) | [English](#english)

---

<a id="中文"></a>

## 中文

Windows 11 资源管理器多窗口自动合并为单窗口多标签页的 AutoHotkey v2 脚本（[1explorer.ahk](1explorer.ahk)）。

基于 Reddit 上的[原始版本](https://www.reddit.com/r/AutoHotkey/comments/1fvsm2y/force_windows_11_to_open_file_explorer_in_new_tab/)，针对中文系统和控制面板窗口做了大量修正。

### 用法

1. 安装 [AutoHotkey v2](https://www.autohotkey.com/)。
2. 双击 `1explorer.ahk` 运行；如需开机自启，放入 `shell:startup` 文件夹。
3. 运行后：新打开的资源管理器窗口会自动合并为第一个窗口的标签页。

### 调试模式

排查问题（如某个窗口没被合并）时，可开启调试日志：

- **开关**：右键托盘图标 → **调试模式** 切换；或带参数启动 `AutoHotkey64.exe 1explorer.ahk --debug`（从启动即记录）。
- **输出**：追加写入脚本同目录的 `1explorer_debug.log`，同时输出到 `OutputDebug`（可用 DebugView 实时查看）。
- **内容**：每个窗口的隐藏/合并/跳过及原因、各步失败点；某页因"导航弹出新窗口"被判定不可合并时，日志会打出完整路径，末尾 CLSID 可直接加入 `knownBad`；反之 `宽限期无弹窗` 确认可合并的页可加入 `knownGood` 实现直通。

### 工作原理

通过 `SetWinEventHook` 监听 `CabinetWClass`（资源管理器窗口）的系统事件：

| 事件 | 处理 |
|---|---|
| `EVENT_OBJECT_CREATE` | 尽早把新窗口设为全透明（视觉隐藏），避免闪烁 |
| `EVENT_OBJECT_SHOW` | 第一个窗口作为主窗口 `FirstWindow`；其余窗口延迟进入 `MergeWindow` 合并流程 |
| `EVENT_OBJECT_DESTROY` | 清理忽略列表 |

`MergeWindow` 合并流程：

1. `GetPath` 通过 `Shell.Application.Windows` + `IShellBrowser` 取得窗口当前路径（带重试，等 COM 就绪）。取到的路径先做格式校验（合法虚拟路径或真实存在的文件系统路径）：`ShellWindows` 里残留的已关闭窗口对象会读回已释放内存中的垃圾串（如 `OleMainThreadWndClass A2DD-…}\::{…}`），校验不过则继续重试，超时按"取不到路径"恢复窗口。
2. `IsExcludedPath` 判断该路径是否不可合并（见下节）。
3. `OpenInNewTab` 向主窗口发送 `WM_COMMAND 0xA21B` 新建标签页，再用 `IWebBrowser2.Navigate2` 导航到目标路径：
   - 普通文件路径 → 直接 `Navigate2(path)`；
   - 虚拟文件夹（`::{GUID}`，含 `::{父}\::{子}` 复合路径）→ `SHParseDisplayName` 解析为 PIDL，包装成 `VT_ARRAY|VT_UI1` 的 VARIANT 后 `Navigate2`。
4. **弹窗检测**：导航后等 1 秒对比窗口列表快照。某些页（如"管理存储空间"）无法原地导航，系统会另外弹出一个独立窗口——检测到新窗口即判定该页不可合并。
5. **宽限确认**：合并看似成功也不立即关闭原窗口，而是等 2 秒（`PendingMerges` + `FinalizeMerge`）。期间若有同路径新窗口迟到弹出，则关闭弹出的复制品、恢复**原来的 Explorer 窗口**（避免窗口"跑到控制面板"图标下）；确认无弹窗才关闭原窗口完成合并。宽限期确认无弹窗的路径记入 `GoodPaths`，本次运行内再次合并同路径时**跳过 1 秒弹窗检测和宽限期**，立即关闭原窗口；关后仍保留 2 秒记录，期间同路径迟到的弹窗复制品会被直接关闭。

任何一步失败（取不到路径、解析失败、导航失败、弹出新窗口）都会恢复原窗口显示并记入忽略列表——**不报错、不丢窗口**。检测到弹窗时还会关闭弹出的复制品，并把主窗口里新建的标签页导航回导航前位置，撤销"整个窗口被挪到控制面板任务栏分组"的副作用。

### 控制面板页的处理策略（四层）

判断一律基于**路径**而非窗口文字——窗口文字在 `EVENT_OBJECT_SHOW` 时尚未填充，存在时序竞争；且中文系统显示文本不同。附带好处：文件夹名称里含"控制面板"也能正常合并。

1. **控制面板类别视图整族 / 所有控制面板项**：按**路径前缀**排除，不再逐页匹配——凡以 `::{26EE0668-…}` 开头的路径（主页、`::{26EE0668-…}\0`、`…\5\::{BitLocker}`、`…\0\::{管理存储空间}` 等）一律在**导航之前**跳过，窗口保持独立；"所有控制面板项"`::{21EC2020-…}`（含复合路径结尾）同样排除。判断发生在 `OpenInNewTab` 之前，所以不会出现"主窗口被挪到控制面板任务栏分组"。`knownGood` 白名单优先于本规则。
2. **已知"导航会弹新窗口"的页**：按**末尾 CLSID** 匹配直接跳过（入口不同，复合路径的父 GUID 不同，故不匹配完整路径）。目前名单（`knownBad`）：
   - `{F942C606-0914-47AB-BE56-1321B8035096}` 管理存储空间
   - `{D9EF8727-CAC2-4E60-809E-86F80A666C91}` BitLocker 驱动器加密
3. **已知可合并的页**：同样按**末尾 CLSID** 匹配，跳过 1 秒弹窗检测和 2 秒宽限期，导航成功后立即关闭原窗口。目前名单（`knownGood`）：
   - `{7007ACC7-3202-11D1-AAD2-00805FC1270E}` 网络和共享中心
4. **未知弹窗页**：运行时自动识别（第 4 步弹窗检测），记入 `BadPaths`，本次运行内再次打开时直接跳过。

#### 如何添加新的已知弹窗页

发现某个控制面板页不能合并时，先保持该窗口打开，用 PowerShell 查询其路径（或者开启[调试模式](#中文)后打开该页，日志里会直接打出完整路径）：

```powershell
(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Document.Folder.Self.Path }
```

取输出的最后一段 CLSID，加到 `IsExcludedPath` 中 `knownBad` 的 `(?:...)` 列表里即可（本次运行已自动识别的页重启脚本后仍需加入名单才会直接跳过）。反过来，确认**能合并**的页（日志里出现过 `宽限期无弹窗`）可把末尾 CLSID 加到 `IsKnownGoodPath` 的 `knownGood` 列表，之后任何入口打开都直接跳过检测与宽限。

控制面板项 CLSID 与名称对照可用以下命令查询：

```powershell
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace' | ForEach-Object { $guid = $_.PSChildName; $name = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\CLSID\$guid" -ErrorAction SilentlyContinue).'(default)'; "$guid  $name" }
```

### 已知限制

- 首次遇到某个未知弹窗页时，主窗口会多出一个标签页：检测失败后会自动把它导航回导航前的"主页"，因此不会把主窗口"挪到控制面板"任务栏分组，但这个多余标签页仍需手动关闭；第二次打开该页起不再出现。
- `BadPaths`（运行时自动识别的名单）只在本次运行有效，重启脚本后需重新识别一次；常遇到的页建议按上节方法加入 `knownBad`。
- `GoodPaths`（已知可合并、跳过弹窗检测与宽限期的名单）同样只在本次运行有效。若某路径极偶尔才会迟到弹窗，跳过检测后来不及拦截，弹出的复制品可能让主窗口多出一个重复标签页。
- 合并成功后的 2 秒宽限期内，如果用户手动再次打开**同一路径**的窗口，可能被误判为"弹窗"而关闭并把该路径标记为不可合并（极少见，重启脚本恢复）。
- 若第一个打开的窗口本身就是控制面板主页，`CheckFirstWindow` 会将其降级为独立窗口，后续文件夹窗口另立主窗口。

### 关键实现

- 隐藏方式用 `WinSetTransparent(0)` 而非隐藏窗口，规避 `DetectHiddenWindows` 的干扰。
- `WaitForSameWindowCount` 等待 `Shell.Application.Windows` 数量与标签控件数量一致（忽略已排除窗口），保证 `GetPath`/`OpenInNewTab` 按索引取到正确的 shell 窗口。
- `MergeLock` 保证一次只合并一个窗口——弹窗检测依赖窗口列表快照，并发合并会互相干扰。等锁 15 秒超时则放弃本次合并并恢复原窗口（不无锁硬闯，避免连环误判）。
- 原窗口关闭前始终处于透明隐藏状态，因此 1 秒弹窗检测和 2 秒宽限期用户均无感知。

---

<a id="english"></a>

## English

An AutoHotkey v2 script that merges multiple Windows 11 File Explorer windows into a single window with multiple tabs ([1explorer.ahk](1explorer.ahk)).

Based on the [original version](https://www.reddit.com/r/AutoHotkey/comments/1fvsm2y/force_windows_11_to_open_file_explorer_in_new_tab/) from Reddit, with extensive fixes for non-English (Chinese) systems and Control Panel windows.

### Usage

1. Install [AutoHotkey v2](https://www.autohotkey.com/).
2. Double-click `1explorer.ahk` to run it. To start it automatically at login, place it in your `shell:startup` folder.
3. Once running: newly opened File Explorer windows are automatically merged as tabs into the first window.

### Debug Mode

To troubleshoot (e.g. a window that won't merge), enable debug logging:

- **Toggle**: right-click the tray icon → **调试模式**; or start with `AutoHotkey64.exe 1explorer.ahk --debug` to log from startup.
- **Output**: appended to `1explorer_debug.log` next to the script, and also sent to `OutputDebug` (viewable live with DebugView).
- **Content**: every window hide / merge / skip decision with its reason, plus failure points. When a page is rejected because navigation spawned a new window, the log prints its full path — copy the trailing CLSID into `knownBad`; conversely, pages confirmed mergeable by `宽限期无弹窗` can be added to `knownGood` for a fast path.

### How It Works

It listens for system events on `CabinetWClass` (File Explorer windows) via `SetWinEventHook`:

| Event | Handling |
|---|---|
| `EVENT_OBJECT_CREATE` | Makes the new window fully transparent (visually hidden) as early as possible to avoid flicker |
| `EVENT_OBJECT_SHOW` | The first window becomes the main window `FirstWindow`; other windows enter the `MergeWindow` flow after a delay |
| `EVENT_OBJECT_DESTROY` | Cleans up the ignore list |

The `MergeWindow` flow:

1. `GetPath` obtains the window's current path via `Shell.Application.Windows` + `IShellBrowser` (with retries while waiting for COM to become ready). The path is then format-validated (a well-formed virtual path or an existing file-system path): stale objects of already-closed windows left in `ShellWindows` can return garbage read from freed memory (e.g. `OleMainThreadWndClass A2DD-…}\::{…}`). If validation fails it keeps retrying, and on timeout the window is restored as "path unavailable".
2. `IsExcludedPath` decides whether the path cannot be merged (see the next section).
3. `OpenInNewTab` sends `WM_COMMAND 0xA21B` to the main window to open a new tab, then navigates it to the target path via `IWebBrowser2.Navigate2`:
   - Regular file paths → `Navigate2(path)` directly;
   - Virtual folders (`::{GUID}`, including composite paths like `::{parent}\::{child}`) → `SHParseDisplayName` resolves them into a PIDL, which is wrapped into a `VT_ARRAY|VT_UI1` VARIANT before calling `Navigate2`.
4. **Popup detection**: after navigating, wait 1 second and compare the window-list snapshot. Some pages (such as "Manage Storage Spaces") cannot be navigated in place — Windows instead pops up a separate window. Any new window is treated as proof that the page cannot be merged.
5. **Grace-period confirmation**: even when the merge appears to succeed, the original window is not closed immediately. It waits 2 seconds (`PendingMerges` + `FinalizeMerge`). If a new window with the same path appears late, the duplicate is closed and the **original Explorer window** is restored (so the window never "jumps" under the Control Panel icon); the original window is closed only after the grace period confirms no popup. Paths that pass the grace period are recorded in `GoodPaths`; merging the same path again within the session **skips both the 1-second popup check and the grace period**, closing the original window immediately. A 2-second record is still kept afterwards so that a late popup duplicate of the same path is closed right away.

If any step fails (path unavailable, parse failure, navigation failure, popup detected), the original window is restored and added to the ignore list — **no errors, no lost windows**. When a popup is detected, the duplicate is closed and the newly created tab in the main window is navigated back to its previous location, undoing the "whole window moves under the Control Panel taskbar group" side effect.

### Control Panel Pages — Four-Tier Strategy

All decisions are based on **paths**, never on window text. Window text is not yet populated when `EVENT_OBJECT_SHOW` fires (a race condition), and the displayed text differs on Chinese systems. A side benefit: folders whose names contain "Control Panel" now merge normally.

1. **Control Panel category-view family / All Control Panel Items**: excluded by **path prefix** instead of page-by-page matching — anything starting with `::{26EE0668-…}` (the home page, `::{26EE0668-…}\0`, `…\5\::{BitLocker}`, `…\0\::{Storage Spaces}`, etc.) is skipped **before navigation**, so the window stays standalone; "All Control Panel Items" `::{21EC2020-…}` (including composite-path endings) is excluded the same way. The check runs before `OpenInNewTab`, so the "main window moves under the Control Panel taskbar group" side effect cannot happen. The `knownGood` whitelist takes precedence over this rule.
2. **Known pages that pop up a new window**: matched by their **trailing CLSID** and skipped directly (the parent GUID in a composite path differs depending on the entry point, so full-path matching is not used). Current list (`knownBad`):
   - `{F942C606-0914-47AB-BE56-1321B8035096}` Manage Storage Spaces
   - `{D9EF8727-CAC2-4E60-809E-86F80A666C91}` BitLocker Drive Encryption
3. **Known mergeable pages**: also matched by their **trailing CLSID**; they skip the 1-second popup check and the 2-second grace period, and the original window is closed immediately after navigation. Current list (`knownGood`):
   - `{7007ACC7-3202-11D1-AAD2-00805FC1270E}` Network and Sharing Center
4. **Unknown popup pages**: detected automatically at runtime (step 4 above) and recorded in `BadPaths`; opening the same page again within the current session skips it directly.

#### Adding a New Known Popup Page

When you find a Control Panel page that cannot be merged, keep that window open and query its path with PowerShell (or enable [debug mode](#english) and open the page — the log prints its full path directly):

```powershell
(New-Object -ComObject Shell.Application).Windows() | ForEach-Object { $_.Document.Folder.Self.Path }
```

Take the last CLSID segment from the output and add it to the `(?:...)` list of `knownBad` in `IsExcludedPath`. (Pages auto-detected during a run still need to be added to this list to be skipped instantly after a restart.) Conversely, pages confirmed **mergeable** (`宽限期无弹窗` in the log) can have their trailing CLSID added to the `knownGood` list in `IsKnownGoodPath`, after which they always skip the popup check and the grace period regardless of entry point.

To look up Control Panel item CLSIDs and their names:

```powershell
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace' | ForEach-Object { $guid = $_.PSChildName; $name = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\CLSID\$guid" -ErrorAction SilentlyContinue).'(default)'; "$guid  $name" }
```

### Known Limitations

- The first time an unknown popup page is encountered, the main window gains an extra tab: after the failure is detected it is navigated back to the previous "Home", so the main window does not move under the Control Panel taskbar group — but the extra tab still has to be closed manually; from the second time onward it no longer appears.
- `BadPaths` (the runtime auto-detected list) is only valid for the current session. After restarting the script the page is detected again; pages you encounter often should be added to `knownBad` as described above.
- `GoodPaths` (the known-mergeable list that skips the popup check and the grace period) is also session-only. If a path very rarely pops up a window late, skipping the check means the popup may no longer be intercepted in time, and the main window may gain an extra duplicate tab.
- During the 2-second grace period after a successful merge, manually opening another window for the **same path** may be mistaken for a "popup" — that window is closed and the path marked as unmergeable (very rare; restarting the script clears it).
- If the very first window opened is the Control Panel home page, `CheckFirstWindow` downgrades it to a standalone window and the next folder window becomes the new main window.

### Implementation Notes

- Windows are hidden with `WinSetTransparent(0)` instead of being hidden, to avoid interference with `DetectHiddenWindows`.
- `WaitForSameWindowCount` waits until the `Shell.Application.Windows` count matches the number of tab controls (ignoring excluded windows), so that `GetPath`/`OpenInNewTab` index into the correct shell window.
- `MergeLock` ensures only one window is merged at a time — popup detection relies on window-list snapshots, and concurrent merges would interfere with each other. Waiting for the lock times out after 15 seconds: the merge is then abandoned and the original window restored (never merging without the lock, which would cascade into false popup detections).
- The original window always stays transparently hidden before it is closed, so the 1-second popup detection and the 2-second grace period are invisible to the user.

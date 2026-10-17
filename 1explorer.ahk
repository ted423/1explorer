#Requires AutoHotkey v2.0
; https://www.reddit.com/r/AutoHotkey/comments/1fvsm2y/force_windows_11_to_open_file_explorer_in_new_tab/
; 系统中文，控制面板显示不一样，改为按路径判断是否合并，不再依赖窗口文字（解决时序竞争）：
; - 控制面板主页 / 所有控制面板项：按末尾 CLSID 排除（兼容复合路径形式）
; - 已知"导航会弹新窗口"的页（如 管理存储空间）：按末尾 CLSID 直接跳过（knownBad 名单）
; - 已知可合并的页（如 网络和共享中心）：按末尾 CLSID 直通，跳过弹窗检测与宽限期（knownGood 名单）
; - 未知弹窗页：运行时自动识别（Navigate2 后检测新窗口），记入 BadPaths 本次运行内不再尝试
; - 已成功合并（宽限期确认无弹窗）的路径记入 GoodPaths，本次运行内再次合并时跳过 1 秒弹窗检测与宽限期，立即关原窗口
; - 路径先校验格式再使用，规避 ShellWindows 残留对象返回的垃圾串（如 "OleMainThreadWndClass ..."）
; - 读取路径或导航失败时恢复窗口显示，不再报错、不丢窗口
; - 附带修复：文件夹名称里含"控制面板"也能正常合并了
; 调试：右键托盘图标切"调试模式"，或带 --debug 参数启动；日志写入脚本同目录 1explorer_debug.log
Persistent

ForceOneExplorerWindow()

class ForceOneExplorerWindow {

    static __New() {
        this.FirstWindow := 0
        this.hHook := 0
        this.pWinEventHook := CallbackCreate(ObjBindMethod(this, 'WinEventProc'),, 7)
        this.IgnoreWindows := Map()
        this.BadPaths := Map()      ; 实测无法在标签页打开的页（导航会弹新窗口），本次运行内不再尝试
        this.GoodPaths := Map()     ; 宽限期确认无弹窗、已成功合并过的路径，再次合并时跳过宽限期
        this.PendingMerges := Map() ; path -> 等待"未弹窗"确认的原窗口 hwnd（宽限 2 秒后才关闭）
        this.MergeLock := false
        this.shellWindows := ComObject('Shell.Application').Windows
        this.DebugLog := A_ScriptDir '\1explorer_debug.log'
        this.DebugMode := false
        for arg in A_Args {
            if arg = '--debug' || arg = '/debug'
                this.DebugMode := true
        }
        A_TrayMenu.Add()
        A_TrayMenu.Add('调试模式', ObjBindMethod(this, 'ToggleDebug'))
        if this.DebugMode
            A_TrayMenu.Check('调试模式')
        this.Log('脚本启动')
    }

    static Call() {
        this.MergeWindows()
        if !this.hHook {
            this.hHook := DllCall('SetWinEventHook', 'uint', 0x8000, 'uint', 0x8002, 'ptr', 0, 'ptr', this.pWinEventHook
                                , 'uint', 0, 'uint', 0, 'uint', 0x2, 'ptr')
        }
    }

    static GetPath(hwnd) {
        static IID_IShellBrowser := '{000214E2-0000-0000-C000-000000000046}'
        shellWindows := this.shellWindows
        lastBad := ''
        this.WaitForSameWindowCount()
        try activeTab := ControlGetHwnd('ShellTabWindowClass1', hwnd)
        loop 20 {  ; 等待新窗口的 COM 对象就绪，最多约 2 秒
            for w in shellWindows {
                if w.hwnd != hwnd
                    continue
                path := ''
                try {
                    if IsSet(activeTab) {
                        shellBrowser := ComObjQuery(w, IID_IShellBrowser, IID_IShellBrowser)
                        ComCall(3, shellBrowser, 'uint*', &thisTab:=0)
                        if thisTab = activeTab
                            path := w.Document.Folder.Self.Path
                    } else {
                        path := w.Document.Folder.Self.Path
                    }
                }
                if this.IsValidPath(path)
                    return path
                if path != ''
                    lastBad := path
            }
            sleep 100
        }
        this.Log('GetPath 超时，取不到有效路径 hwnd=' hwnd '，最后读到：' lastBad)
        return ''
    }

    ; ShellWindows 里可能残留刚关闭窗口的对象：此时读 Document.Folder.Self.Path 会得到已释放内存里的垃圾串
    ; （如 "OleMainThreadWndClass A2DD-08002B30309D}\::{...}"），必须校验后再用
    static IsValidPath(path) {
        if path = ''
            return false
        ; 虚拟路径（含 ::{父}\::{子} 复合路径）；具体是否存在交给后续 SHParseDisplayName 校验
        if path ~= 'i)^::\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}'
            return true
        return FileExist(path) != ''
    }

    ; 只排除无法作为标签页打开的页（一律按 CLSID 匹配）：
    ; 同一页从不同入口打开会返回复合路径（如 ::{26EE0668-…}\0\::{F942C606-…}），末尾 CLSID 才是当前页
    ; - 控制面板主页 / 所有控制面板项：合并会把主窗口整个挪到"控制面板"任务栏图标下
    ;   且类别视图会返回 ::{26EE0668-…}\0 这种"主页+数字段"形式（末尾不是 CLSID，需单独匹配）
    ; - 已知"导航会弹新窗口"的页
    static IsExcludedPath(path) {
        static knownBad := '(?:F942C606-0914-47AB-BE56-1321B8035096)'  ; F942C606 = 管理存储空间
        static home := '(?:26EE0668-A00A-44D7-9371-BEB064C98683|21EC2020-3AEA-1069-A2DD-08002B30309D)'
        return path ~= 'i)::\{' home '\}$'              ; 末尾 CLSID 是主页（含复合路径形式）
            || path ~= 'i)^::\{' home '\}(?:\\[^\\]*)?$'  ; 主页 + 一段数字/类别（如 ::{26EE0668-…}\0）
            || path ~= 'i)::\{' knownBad '\}$'
    }

    ; 已知可合并的页（按末尾 CLSID 匹配，兼容不同入口的复合路径）：
    ; 跳过弹窗检测与宽限期，导航成功后立即关闭原窗口
    static IsKnownGoodPath(path) {
        static knownGood := '(?:7007ACC7-3202-11D1-AAD2-00805FC1270E)'  ; 7007ACC7 = 网络和共享中心
        return path ~= 'i)::\{' knownGood '\}$'
    }

    static MergeWindows() {
        windows := WinGetList('ahk_class CabinetWClass')
        if windows.Length = 0
            return
        this.FirstWindow := windows.RemoveAt(1)
        this.Log('启动：主窗口 hwnd=' this.FirstWindow)
        if WinGetTransparent(this.FirstWindow) = 0 {
            WinSetTransparent("Off", this.FirstWindow)
        }
        firstWindow := this.FirstWindow
        shellWindows := this.shellWindows
        seen := Map()
        keep := Map()
        paths := []
        for w in shellWindows {
            if w.hwnd = firstWindow
                continue
            seen.Set(w.hwnd, 1)
            path := ''
            try path := w.Document.Folder.Self.Path
            if !this.IsValidPath(path) || this.IsExcludedPath(path) {
                keep.Set(w.hwnd, 1)
                this.Log('启动合并：保留独立窗口 hwnd=' w.hwnd ' path=' path)
                continue
            }
            paths.push(path)
        }
        for hwnd in windows {
            if !seen.Has(hwnd) || keep.Has(hwnd) {
                if WinGetTransparent(hwnd) = 0  ; 上次异常退出可能残留透明窗口，恢复显示
                    WinSetTransparent('Off', hwnd)
                continue  ; 读不到路径或不需要合并的窗口保持原样
            }
            PostMessage(0x0112, 0xF060,,, hwnd)  ; 0x0112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
            WinWaitClose(hwnd)
        }
        for path in paths {
            this.Log('启动合并：新标签页打开 ' path)
            this.OpenInNewTab(path)
        }
    }

    static WinEventProc(hWinEventHook, event, hwnd, idObject, idChild, idEventThread, dwmsEventTime) {
        Critical(-1)
        if !(idObject = 0 && idChild = 0) {
            return
        }
        switch event {
            case 0x8000:  ; EVENT_OBJECT_CREATE
                ancestor := DllCall('GetAncestor', 'ptr', hwnd, 'uint', 2, 'ptr')
                try {
                    if !this.IgnoreWindows.Has(ancestor) && WinExist(ancestor) && WinGetClass(ancestor) = 'CabinetWClass' {
                        if ancestor = this.FirstWindow
                            return
                        if WinGetTransparent(ancestor) = '' {
                            ; Hide window as early as possible
                            WinSetTransparent(0, ancestor)
                            this.Log('CREATE：隐藏新窗口 hwnd=' ancestor)
                        }
                    }
                }
            case 0x8002:  ; EVENT_OBJECT_SHOW
                if WinExist(hwnd) && WinGetClass(hwnd) = 'CabinetWClass' {
                    if this.IgnoreWindows.Has(hwnd)
                        return
                    if !WinExist(this.FirstWindow) {
                        this.FirstWindow := hwnd
                        WinSetTransparent('Off', hwnd)
                        this.Log('SHOW：设为主窗口 hwnd=' hwnd)
                        SetTimer(() => this.CheckFirstWindow(hwnd), -1)
                        return
                    }
                    if WinGetTransparent(hwnd) = 0 {
                        this.Log('SHOW：进入合并流程 hwnd=' hwnd)
                        SetTimer(() => this.MergeWindow(hwnd), -1)
                    }
                }
            case 0x8001:  ; EVENT_OBJECT_DESTROY
                if this.IgnoreWindows.Has(hwnd) {
                    this.IgnoreWindows.Delete(hwnd)
                    this.Log('DESTROY：清理忽略列表 hwnd=' hwnd)
                }
        }
    }

    ; 把 hwnd 合并为 FirstWindow 的标签页；无法合并时恢复显示并忽略该窗口
    static MergeWindow(hwnd) {
        timeout := A_TickCount + 15000
        while this.MergeLock {  ; 一次只合并一个，弹出窗口检测依赖窗口列表快照
            sleep 50
            if A_TickCount > timeout {
                ; 上一个合并被卡住（如恰逢另一窗口创建，计数等待吃满超时），放弃本次合并而非无锁硬闯
                this.Log('等合并锁超时，放弃合并并恢复窗口 hwnd=' hwnd)
                this.RestoreWindow(hwnd)
                return
            }
        }
        if !WinExist(hwnd)
            return
        this.Log('开始合并 hwnd=' hwnd)
        path := this.GetPath(hwnd)
        if path = '' || this.IsExcludedPath(path) || this.BadPaths.Has(path) {
            reason := path = '' ? '取不到路径' : this.IsExcludedPath(path) ? '排除列表' : 'BadPaths'
            this.Log('跳过合并（' reason '）hwnd=' hwnd ' path=' path)
            this.RestoreWindow(hwnd)
            return
        }
        if this.PendingMerges.Has(path) {
            orig := this.PendingMerges[path]
            this.PendingMerges.Delete(path)
            if orig = 0 {
                ; 已知可合并路径刚合并完（原窗口已提前关闭）：这是迟到弹窗的复制品，直接关掉
                this.Log('迟到弹窗：已知可合并路径，关闭复制品 hwnd=' hwnd ' path=' path)
                if WinExist(hwnd)
                    WinClose(hwnd)
                return
            }
            ; 上一次合并的页弹出了新窗口：该页无法在标签页原地打开
            ; 关掉弹出的复制品，恢复原来的 Explorer 窗口（避免窗口"跑到控制面板"）
            this.BadPaths[path] := 1
            if this.GoodPaths.Has(path)  ; Map.Delete 键不存在会抛异常
                this.GoodPaths.Delete(path)
            this.Log('迟到弹窗：关闭复制品 hwnd=' hwnd '，恢复原窗口 hwnd=' orig '；判定不可合并 path=' path '（可把末尾 CLSID 加入 knownBad）')
            if WinExist(hwnd)
                WinClose(hwnd)
            this.RestoreWindow(orig)
            return
        }
        spawned := 0
        ok := false
        ; 已确认可合并（运行时验证过或静态 knownGood 名单）：跳过 1 秒弹窗检测和宽限期，立即关原窗口
        isGood := this.GoodPaths.Has(path) || this.IsKnownGoodPath(path)
        this.MergeLock := true
        try {
            before := Map()
            for h in WinGetList('ahk_class CabinetWClass')
                before[h] := 1
            ok := this.OpenInNewTab(path)
            if ok && !isGood {
                sleep 1000  ; 等可能弹出的新窗口（原窗口此时是隐藏的，用户无感知）
                for h in WinGetList('ahk_class CabinetWClass') {
                    if !before.Has(h) && h != hwnd {
                        spawned := h
                        break
                    }
                }
            }
        } finally {
            this.MergeLock := false
        }
        if ok && !spawned {
            if isGood {
                this.Log('合并成功（已知可合并，立即关闭原窗口）hwnd=' hwnd ' path=' path)
                if WinExist(hwnd)
                    WinClose(hwnd)
                ; 原窗口已关，但仍留 2 秒记录：期间同路径迟到弹窗的复制品直接关掉（无需恢复）
                this.PendingMerges[path] := 0
                SetTimer(ObjBindMethod(this, 'FinalizeMerge', path), -2000)
            } else {
                ; 暂不关闭原窗口：宽限 2 秒确认没有迟到的弹窗后，由 FinalizeMerge 关闭
                this.PendingMerges[path] := hwnd
                SetTimer(ObjBindMethod(this, 'FinalizeMerge', path), -2000)
                this.Log('合并成功，宽限 2 秒后关闭原窗口 hwnd=' hwnd ' path=' path)
            }
            if WinExist(this.FirstWindow) && WinGetMinMax(this.FirstWindow) = -1
                WinRestore(this.FirstWindow)
        } else {
            if spawned {
                this.BadPaths[path] := 1
                if this.GoodPaths.Has(path)  ; Map.Delete 键不存在会抛异常
                    this.GoodPaths.Delete(path)
                this.Log('导航弹出新窗口 hwnd=' spawned '；判定不可合并 path=' path '（可把末尾 CLSID 加入 knownBad）')
                if WinExist(spawned)
                    WinClose(spawned)
            } else {
                this.Log('OpenInNewTab 失败 path=' path)
            }
            this.RestoreWindow(hwnd)
        }
    }

    static RestoreWindow(hwnd) {
        this.IgnoreWindows.Set(hwnd, 1)
        this.Log('恢复窗口显示并忽略 hwnd=' hwnd)
        if WinExist(hwnd)
            WinSetTransparent('Off', hwnd)
    }

    ; 宽限期结束仍没弹窗：合并确认成功，关闭原窗口
    static FinalizeMerge(path) {
        if !this.PendingMerges.Has(path)
            return  ; 弹窗分支已处理
        hwnd := this.PendingMerges[path]
        this.PendingMerges.Delete(path)
        this.GoodPaths[path] := 1  ; 已验证可合并，下次同路径跳过宽限期
        if hwnd {  ; hwnd = 0 表示已知可合并路径：原窗口早已关闭，仅清理记录
            this.Log('宽限期无弹窗，关闭原窗口 hwnd=' hwnd ' path=' path)
            if WinExist(hwnd)
                WinClose(hwnd)
        } else {
            this.Log('宽限期无弹窗（原窗口已提前关闭）path=' path)
        }
    }

    ; 若 FirstWindow 本身是无法合并的页（如控制面板主页），改回独立窗口
    static CheckFirstWindow(hwnd) {
        if hwnd != this.FirstWindow || !WinExist(hwnd)
            return
        path := this.GetPath(hwnd)
        if path = '' || this.IsExcludedPath(path) || this.BadPaths.Has(path) {
            this.IgnoreWindows.Set(hwnd, 1)
            this.Log('主窗口不可合并，降级为独立窗口 hwnd=' hwnd ' path=' path)
            if hwnd = this.FirstWindow
                this.FirstWindow := 0
        }
    }

    ; 调试模式：托盘菜单切换，或带 --debug 参数启动
    static ToggleDebug(*) {
        if this.DebugMode {
            this.Log('=== 调试模式关闭 ===')
            this.DebugMode := false
            A_TrayMenu.Uncheck('调试模式')
        } else {
            this.DebugMode := true
            A_TrayMenu.Check('调试模式')
            this.Log('=== 调试模式开启 ===')
            TrayTip('日志写入 ' this.DebugLog, '1explorer 调试模式已开启')
        }
    }

    ; 调试日志：写脚本同目录的 1explorer_debug.log，同时输出到 OutputDebug（可用 DebugView 查看）
    static Log(msg) {
        if !this.DebugMode
            return
        line := FormatTime(, 'HH:mm:ss') '.' Format('{:03}', A_MSec) ' ' msg
        OutputDebug('1explorer ' line)
        try FileAppend(line '`n', this.DebugLog, 'UTF-8')
    }

    static WaitForSameWindowCount() {
        shellWindows := this.shellWindows
        windowCount := 0
        for hwnd in WinGetList('ahk_class CabinetWClass') {
            if this.IgnoreWindows.Has(hwnd)
                continue
            for classNN in WinGetControls(hwnd) {
                if classNN ~= '^ShellTabWindowClass\d+'
                    windowCount++
            }
        }
        ; wait for window count to update（忽略不合并的窗口，避免每次都等到超时）
        ; 注意必须在循环内实时重算：等的正是 shellWindows 数量变化
        timeout := A_TickCount + 3000
        expected := -1
        while windowCount != expected {
            expected := 0
            for w in shellWindows {
                try {
                    if !this.IgnoreWindows.Has(w.hwnd)
                        expected++
                }
            }
            if windowCount = expected
                break
            sleep 50
            if A_TickCount > timeout {
                this.Log('WaitForSameWindowCount 超时：标签数=' windowCount ' shell窗口数=' expected)
                break
            }
        }
    }

    ; 在 FirstWindow 中为新路径打开标签页，返回是否成功
    static OpenInNewTab(path) {
        shellWindows := this.shellWindows
        ; 先解析路径，解析失败就不创建标签页
        PIDL := 0
        if !FileExist(path) {
            parsePath := path
            ; matches a shell folder path such as ::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}
            if parsePath ~= 'i)^::{[0-9A-F-]+}$'
                parsePath := 'shell:' parsePath
            ; SHParseDisplayName 也支持 ::{guid}\::{guid} 形式的复合路径（如"网络连接"）
            if DllCall('shell32\SHParseDisplayName', 'wstr', parsePath, 'ptr', 0, 'ptr*', &PIDL, 'uint', 0, 'ptr', 0) != 0 || PIDL = 0 {
                this.Log('SHParseDisplayName 解析失败 path=' path)
                return false
            }
        }
        if !WinExist(this.FirstWindow) {
            if PIDL
                DllCall('ole32\CoTaskMemFree', 'ptr', PIDL)
            return false
        }
        this.WaitForSameWindowCount()
        hwnd := this.FirstWindow
        Count := shellWindows.Count()
        ; open a new tab (https://stackoverflow.com/a/78502949)
        SendMessage(0x0111, 0xA21B, 0, 'ShellTabWindowClass1', hwnd)
        ; Wait for window count to change
        timeout := A_TickCount + 3000
        while shellWindows.Count() = Count {
            sleep 50
            if A_TickCount > timeout {
                this.Log('新建标签页超时 path=' path)
                if PIDL
                    DllCall('ole32\CoTaskMemFree', 'ptr', PIDL)
                return false
            }
        }
        try Item := shellWindows.Item(Count)
        if !IsSet(Item) {
            this.Log('取不到新标签页对象 path=' path)
            if PIDL
                DllCall('ole32\CoTaskMemFree', 'ptr', PIDL)
            return false
        }
        if PIDL {
            byteCount := DllCall('shell32\ILGetSize', 'ptr', PIDL, 'uint')
            SAFEARRAY := Buffer(16 + 2 * A_PtrSize, 0)
            NumPut 'ushort', 1, SAFEARRAY, 0  ; cDims
            NumPut 'uint', 1, SAFEARRAY, 4  ; cbElements
            NumPut 'ptr', PIDL, SAFEARRAY, 8 + A_PtrSize  ; pvData
            NumPut 'uint', byteCount, SAFEARRAY, 8 + 2 * A_PtrSize  ; rgsabound[1].cElements
        }
        success := true
        try {
            if PIDL
                Item.Navigate2(ComValue(0x2011, SAFEARRAY.ptr))
            else
                Item.Navigate2(path)
        } catch {
            this.Log('Navigate2 失败 path=' path)
            success := false
        }
        if PIDL
            DllCall('ole32\CoTaskMemFree', 'ptr', PIDL)
        ; 某些页（如"管理存储空间"）导航后 Item 会失效，访问 Busy 会抛 E_FAIL，忽略即可
        try {
            deadline := A_TickCount + 5000
            while Item.Busy {
                sleep 50
                if A_TickCount > deadline
                    break
            }
        }
        return success
    }
}

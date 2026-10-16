#Requires AutoHotkey v2.0
; https://www.reddit.com/r/AutoHotkey/comments/1fvsm2y/force_windows_11_to_open_file_explorer_in_new_tab/
; 系统中文，控制面板显示不一样，改为按路径判断是否合并，不再依赖窗口文字（解决时序竞争）：
; - 控制面板主页 / 所有控制面板项：按完整路径排除
; - 已知"导航会弹新窗口"的页（如 管理存储空间）：按末尾 CLSID 直接跳过
; - 未知弹窗页：运行时自动识别（Navigate2 后检测新窗口），记入 BadPaths 本次运行内不再尝试
; - 读取路径或导航失败时恢复窗口显示，不再报错、不丢窗口
; - 附带修复：文件夹名称里含"控制面板"也能正常合并了
Persistent

ForceOneExplorerWindow()

class ForceOneExplorerWindow {

    static __New() {
        this.FirstWindow := 0
        this.hHook := 0
        this.pWinEventHook := CallbackCreate(ObjBindMethod(this, 'WinEventProc'),, 7)
        this.IgnoreWindows := Map()
        this.BadPaths := Map()      ; 实测无法在标签页打开的页（导航会弹新窗口），本次运行内不再尝试
        this.PendingMerges := Map() ; path -> 等待"未弹窗"确认的原窗口 hwnd（宽限 2 秒后才关闭）
        this.MergeLock := false
        this.shellWindows := ComObject('Shell.Application').Windows
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
                if path != ''
                    return path
            }
            sleep 100
        }
        return ''
    }

    ; 只排除无法作为标签页打开的页：
    ; - 控制面板主页 / 所有控制面板项（完整路径匹配）
    ; - 已知"导航会弹新窗口"的页（按末尾 CLSID 匹配，兼容不同入口的复合路径）
    static IsExcludedPath(path) {
        static knownBad := '(?:F942C606-0914-47AB-BE56-1321B8035096)'  ; F942C606 = 管理存储空间
        return path ~= 'i)^::\{(?:26EE0668-A00A-44D7-9371-BEB064C98683|21EC2020-3AEA-1069-A2DD-08002B30309D)\}$'
            || path ~= 'i)::\{' knownBad '\}$'
    }

    static MergeWindows() {
        windows := WinGetList('ahk_class CabinetWClass')
        if windows.Length = 0
            return
        this.FirstWindow := windows.RemoveAt(1)
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
            if path = '' || this.IsExcludedPath(path) {
                keep.Set(w.hwnd, 1)
                continue
            }
            paths.push(path)
        }
        for hwnd in windows {
            if !seen.Has(hwnd) || keep.Has(hwnd)
                continue  ; 读不到路径或不需要合并的窗口保持原样
            PostMessage(0x0112, 0xF060,,, hwnd)  ; 0x0112 = WM_SYSCOMMAND, 0xF060 = SC_CLOSE
            WinWaitClose(hwnd)
        }
        for path in paths {
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
                        SetTimer(() => this.CheckFirstWindow(hwnd), -1)
                        return
                    }
                    if WinGetTransparent(hwnd) = 0 {
                        SetTimer(() => this.MergeWindow(hwnd), -1)
                    }
                }
            case 0x8001:  ; EVENT_OBJECT_DESTROY
                if this.IgnoreWindows.Has(hwnd)
                    this.IgnoreWindows.Delete(hwnd)
        }
    }

    ; 把 hwnd 合并为 FirstWindow 的标签页；无法合并时恢复显示并忽略该窗口
    static MergeWindow(hwnd) {
        timeout := A_TickCount + 15000
        while this.MergeLock {  ; 一次只合并一个，弹出窗口检测依赖窗口列表快照
            sleep 50
            if A_TickCount > timeout
                break
        }
        if !WinExist(hwnd)
            return
        path := this.GetPath(hwnd)
        if path = '' || this.IsExcludedPath(path) || this.BadPaths.Has(path) {
            this.RestoreWindow(hwnd)
            return
        }
        if this.PendingMerges.Has(path) {
            ; 上一次合并的页弹出了新窗口：该页无法在标签页原地打开
            ; 关掉弹出的复制品，恢复原来的 Explorer 窗口（避免窗口"跑到控制面板"）
            orig := this.PendingMerges[path]
            this.PendingMerges.Delete(path)
            this.BadPaths[path] := 1
            if WinExist(hwnd)
                WinClose(hwnd)
            this.RestoreWindow(orig)
            return
        }
        spawned := 0
        ok := false
        this.MergeLock := true
        try {
            before := Map()
            for h in WinGetList('ahk_class CabinetWClass')
                before[h] := 1
            ok := this.OpenInNewTab(path)
            sleep 1000  ; 等可能弹出的新窗口（原窗口此时是隐藏的，用户无感知）
            for h in WinGetList('ahk_class CabinetWClass') {
                if !before.Has(h) && h != hwnd {
                    spawned := h
                    break
                }
            }
        } finally {
            this.MergeLock := false
        }
        if ok && !spawned {
            ; 暂不关闭原窗口：宽限 2 秒确认没有迟到的弹窗后，由 FinalizeMerge 关闭
            this.PendingMerges[path] := hwnd
            SetTimer(ObjBindMethod(this, 'FinalizeMerge', path), -2000)
            if WinExist(this.FirstWindow) && WinGetMinMax(this.FirstWindow) = -1
                WinRestore(this.FirstWindow)
        } else {
            if spawned {
                this.BadPaths[path] := 1
                if WinExist(spawned)
                    WinClose(spawned)
            }
            this.RestoreWindow(hwnd)
        }
    }

    static RestoreWindow(hwnd) {
        this.IgnoreWindows.Set(hwnd, 1)
        if WinExist(hwnd)
            WinSetTransparent('Off', hwnd)
    }

    ; 宽限期结束仍没弹窗：合并确认成功，关闭原窗口
    static FinalizeMerge(path) {
        if !this.PendingMerges.Has(path)
            return  ; 弹窗分支已处理
        hwnd := this.PendingMerges[path]
        this.PendingMerges.Delete(path)
        if WinExist(hwnd)
            WinClose(hwnd)
    }

    ; 若 FirstWindow 本身是无法合并的页（如控制面板主页），改回独立窗口
    static CheckFirstWindow(hwnd) {
        if hwnd != this.FirstWindow || !WinExist(hwnd)
            return
        path := this.GetPath(hwnd)
        if path = '' || this.IsExcludedPath(path) || this.BadPaths.Has(path) {
            this.IgnoreWindows.Set(hwnd, 1)
            if hwnd = this.FirstWindow
                this.FirstWindow := 0
        }
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
            if A_TickCount > timeout
                break
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
            if DllCall('shell32\SHParseDisplayName', 'wstr', parsePath, 'ptr', 0, 'ptr*', &PIDL, 'uint', 0, 'ptr', 0) != 0 || PIDL = 0
                return false
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
                if PIDL
                    DllCall('ole32\CoTaskMemFree', 'ptr', PIDL)
                return false
            }
        }
        try Item := shellWindows.Item(Count)
        if !IsSet(Item) {
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

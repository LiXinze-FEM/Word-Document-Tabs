#Requires AutoHotkey v2.0
#SingleInstance Force
; Mouse hit-testing and drag positions below use absolute screen pixels.
CoordMode("Mouse", "Screen")
; Word Document Tabs
; Based on the user's working Word_Document_Tabs.ahk.
; A separate, movable Word window switcher. It never changes Word files,
; activation/licensing, or Word window dimensions.
; Ctrl+Alt+D: re-dock to Word.  Ctrl+Alt+W: show/hide.
;
; Design: keep the proven WinGetList -> ListView discovery path.
; Preserve order by HWND rather than Word's changing Z-order.
; Draw animated hover/active colors via NM_CUSTOMDRAW.
; Reorder only after an actual ListView drag; do not activate on a drop.

global ConfigFile := A_AppData "\WordSideTabs.ini"
global PaneVisible := false
global Enabled := true
global LastWord := 0
global ActiveId := 0
global Docs := []
global OrderTitles := []
global Anim := Map()
global HoverRow := 0
global HoverSince := 0
global PreviewFor := 0
global ThumbHandle := 0
global PreviewVisible := false
global PreviewAnchorX := 0.0
global PreviewCurrentX := 0.0
global PreviewAnchorY := 0
global Updating := false
global LastSaved := ""
global HasInitialPosition := false

; Mouse drag/drop state (positions are list insertion slots, not HWND order).
global Dragging := false
global DragSource := 0
global DragInsert := 0
global LineCurrentY := 0.0
global LineTargetY := 0.0
global LastScrollTick := 0
global BlockClicksUntil := 0

stored := IniRead(ConfigFile, "Tabs", "Order", "")
if stored != "" {
    for name in StrSplit(stored, "|") {
        if name != ""
            OrderTitles.Push(name)
    }
}

global Pane := Gui("+AlwaysOnTop +ToolWindow +Resize +MinSize205x210", "Word Document Tabs")
Pane.BackColor := "F7F9FD"
Pane.MarginX := 0
Pane.MarginY := 0
Pane.SetFont("s9", "Microsoft YaHei UI")
Pane.AddText("x12 y10 w270 h26 c234568", "📄  Word Document Tabs")

; Retain the native ListView: robust detection, scrolling and click behavior.
; Full-row highlighting is drawn separately and is not tied to native selection.
global LV := Pane.AddListView("x8 y43 w250 h310 -Hdr -Multi +LV0x20 AltSubmit BackgroundFFFFFF", ["Document"])
LV.ModifyCol(1, 230)
LV.OnEvent("Click", SelectDocument)

; Thin drop-position marker, shown ONLY while dragging (never on hover).
global DropLine := Pane.AddProgress("x12 y44 w240 h3 c4E92E4 BackgroundFFFFFF Hidden", 100)
Pane.OnEvent("Size", PaneSize)
Pane.OnEvent("Close", (*) => ExitApp())
Pane.OnEvent("Escape", (*) => ExitApp())

global Preview := Gui("+AlwaysOnTop +ToolWindow -Caption +Border", "Document preview")
Preview.BackColor := "FFFFFF"
Preview.MarginX := 0
Preview.MarginY := 0
Preview.SetFont("s9", "Microsoft YaHei UI")
global PTitle := Preview.AddText("x12 y10 w296 h24 c223A57", "")
global PFallback := Preview.AddText("x18 y90 w280 h90 Center c78869C"
    , "暂无实时缩略图`n点击标签即可切换到该文档")

; Native ListView notifications, registered using AutoHotkey v2's
; control-specific API so GUI event handling remains intact.
LV.OnNotify(-12, PaintRows)      ; NM_CUSTOMDRAW
LV.OnNotify(-109, DragRows)      ; LVN_BEGINDRAG
; Native capture ensures that a drag finishing outside the list is handled.
OnMessage(0x0202, HandleMouseUp)

Hotkey("^!w", Toggle)
Hotkey("^!d", Redock)
OnExit(OnQuit)
SetTimer(Refresh, 700)
SetTimer(MouseTick, 30)
SetTimer(AnimateTick, 25)
SetTimer(SavePosition, 2000)
Refresh()

PaneSize(guiObj, minMax, width, height) {
    global LV, DropLine, HoverRow
    if minMax = -1
        return
    LV.Move(8, 43, Max(160, width - 16), Max(100, height - 53))
    LV.ModifyCol(1, Max(140, width - 38))
    DropLine.Move(12, 44, Max(140, width - 24), 3)
    HoverRow := 0
    HidePreview()
}

Refresh(*) {
    global Updating, Enabled, PaneVisible, LastWord, ActiveId
    global Pane, Docs, HasInitialPosition, Dragging
    if Updating || !Enabled || Dragging
        return
    Updating := true
    try {
        word := WinActive("ahk_class OpusApp ahk_exe WINWORD.EXE")
        panelFocused := WinActive("ahk_id " Pane.Hwnd)
        if word {
            LastWord := word
            ActiveId := word
        } else if !panelFocused {
            HidePreview()
            if PaneVisible {
                Pane.Hide()
                PaneVisible := false
            }
            return
        }

        if !LastWord || !WinExist("ahk_id " LastWord) {
            HidePreview()
            if PaneVisible {
                Pane.Hide()
                PaneVisible := false
            }
            return
        }

        ; Keep the discovery method from the original, working version.
        found := []
        for hwnd in WinGetList("ahk_exe WINWORD.EXE") {
            try {
                if WinGetClass("ahk_id " hwnd) != "OpusApp"
                    continue
                if !DllCall("IsWindowVisible", "ptr", hwnd, "int")
                    continue
                title := Trim(WinGetTitle("ahk_id " hwnd))
                if title = ""
                    continue
                found.Push({id: hwnd, title: title})
            } catch {
                continue
            }
        }

        ; IMPORTANT: Never rebuild based only on WinGetList's Z-order.
        ReconcileDocuments(found)

        if !PaneVisible {
            if !HasInitialPosition {
                InitialDock()
                HasInitialPosition := true
            } else
                Pane.Show("NA") ; no x/y: never jump on document switching
            PaneVisible := true
        }
    } catch {
        ; A temporarily closing Word window should never clear the UI.
    } finally {
        Updating := false
    }
}

ReconcileDocuments(found) {
    global Docs, Anim, OrderTitles
    next := []
    seen := Map()

    ; First, preserve the exact user-defined ordering of still-open HWNDs.
    for old in Docs {
        for current in found {
            if current.id = old.id {
                next.Push(current) ; update title without changing position
                seen[current.id] := true
                break
            }
        }
    }

    ; Newly opened windows append by default; previously ordered documents
    ; are placed at their saved rank, including when opened one by one.
    newlyAdded := false
    for current in found {
        if !seen.Has(current.id) {
            next.Push(current)
            seen[current.id] := true
            newlyAdded := true
        }
    }
    if newlyAdded && OrderTitles.Length > 0
        next := RestoreSavedOrder(next)

    changed := next.Length != Docs.Length
    if !changed {
        for i, doc in next {
            if doc.id != Docs[i].id || doc.title != Docs[i].title {
                changed := true
                break
            }
        }
    }
    if !changed
        return

    Docs := next
    for doc in Docs {
        if !Anim.Has(doc.id)
            Anim[doc.id] := {hover: 0.0, active: 0.0, flash: 0.0}
    }
    ; Do not purge animation objects during the same refresh; HWNDs may close
    ; while Windows dispatches a paint message.
    DrawDocuments()
}

RestoreSavedOrder(items) {
    global OrderTitles
    if OrderTitles.Length = 0
        return items
    sorted := []
    used := Map()
    for savedName in OrderTitles {
        for doc in items {
            if doc.title = savedName && !used.Has(doc.id) {
                sorted.Push(doc)
                used[doc.id] := true
                break
            }
        }
    }
    for doc in items {
        if !used.Has(doc.id)
            sorted.Push(doc)
    }
    return sorted
}

DrawDocuments() {
    global LV, Docs, HoverRow
    HoverRow := 0
    HidePreview()
    LV.Opt("-Redraw")
    try {
        LV.Delete()
        for doc in Docs
            LV.Add("", "📄  " CleanTitle(doc.title))
        LV.ModifyCol(1, Max(150, CurrentListWidth() - 24))
    } finally {
        LV.Opt("+Redraw")
    }
    RepaintList()
}

CurrentListWidth() {
    global LV
    LV.GetPos(, , &width)
    return width
}

CleanTitle(title) {
    return RegExReplace(title, " - (Word|Microsoft Word)$")
}

SelectDocument(ctrl, row) {
    global Docs, ActiveId, LastWord, BlockClicksUntil, Dragging, LV
    if Dragging || A_TickCount < BlockClicksUntil
        return
    if row < 1 || row > Docs.Length
        return

    hwnd := Docs[row].id
    if !WinExist("ahk_id " hwnd)
        return

    ; Native selected state is not the active-document state. The animated
    ; active highlight is owned by ActiveId, so it survives mouse-leave.
    LV.Modify(0, "-Select")
    HidePreview()
    try {
        if WinGetMinMax("ahk_id " hwnd) = -1
            WinRestore("ahk_id " hwnd)
        WinActivate("ahk_id " hwnd)
        ActiveId := hwnd
        LastWord := hwnd
        RepaintList()
        ; Do not call DrawDocuments(), Pane.Show() or Redock() here.
    } catch {
        ; Closed Word windows are removed by the following Refresh tick.
    }
}

; Native drag starts only after the mouse actually moves across a row.
DragRows(ctrl, lParam) {
    global Docs
    row := NumGet(lParam, A_PtrSize = 8 ? 24 : 12, "Int") + 1
    if row >= 1 && row <= Docs.Length
        BeginDrag(row)
}

; Windows' own ListView paints the text; we supply animated row colors.
PaintRows(ctrl, lParam) {
    global Docs, Anim, Dragging, DragSource
    stage := NumGet(lParam, A_PtrSize = 8 ? 24 : 12, "UInt")
    if stage = 0x1 ; CDDS_PREPAINT
        return 0x20 ; CDRF_NOTIFYITEMDRAW

    if stage = 0x10001 { ; CDDS_ITEMPREPAINT
        index := NumGet(lParam, A_PtrSize = 8 ? 56 : 36, "UPtr") + 1
        if index < 1 || index > Docs.Length
            return 0

        doc := Docs[index]
        if Anim.Has(doc.id)
            data := Anim[doc.id]
        else
            data := {hover: 0.0, active: 0.0, flash: 0.0}

        ; Whole-row light-blue color smoothly follows hover + active state.
        bg := MixRGB(0xFFFFFF, 0xD9E9FF, data.active * 0.94)
        bg := MixRGB(bg, 0xE3F0FF, data.hover * 0.80)
        if Dragging && index = DragSource
            bg := MixRGB(bg, 0xBBD7FA, 0.48)
        bg := MixRGB(bg, 0xB9DBFF, data.flash * 0.65)

        foreground := data.active > 0.45 ? 0x164E91 : 0x233B54
        offset := A_PtrSize = 8 ? 80 : 48
        NumPut("UInt", ColorRef(foreground), lParam, offset)
        NumPut("UInt", ColorRef(bg), lParam, offset + 4)
        return 0x2 ; CDRF_NEWFONT: use our item colors
    }
}

BeginDrag(row) {
    global Dragging, DragSource, DragInsert, LV, HoverRow, DropLine
    global LineCurrentY, LineTargetY, LastScrollTick
    Dragging := true
    DragSource := row
    DragInsert := row
    HoverRow := 0
    LastScrollTick := 0
    HidePreview()
    ; The drop marker is drawn over the ListView as a separate child.
    DllCall("SetWindowPos", "ptr", DropLine.Hwnd, "ptr", 0
        , "int", 0, "int", 0, "int", 0, "int", 0, "uint", 0x13)
    DllCall("SetCapture", "ptr", LV.Hwnd, "ptr")
    UpdateDropPosition(true)
    RepaintList()
}

; Sent to the ListView even after the pointer leaves it because of capture.
HandleMouseUp(wParam, lParam, msg, hwnd) {
    global Dragging, LV
    if Dragging && hwnd = LV.Hwnd {
        FinishDrag()
        return 0
    }
}

UpdateDropPosition(initial := false) {
    global Docs, LV, DropLine, DragInsert, Dragging
    global LineTargetY, LineCurrentY, LastScrollTick
    if !Dragging || Docs.Length = 0
        return
    MouseGetPos(&mx, &my)

    pt := Buffer(8, 0)
    NumPut("Int", mx, pt, 0)
    NumPut("Int", my, pt, 4)
    DllCall("ScreenToClient", "ptr", LV.Hwnd, "ptr", pt.Ptr)
    mouseY := NumGet(pt, 4, "Int")

    ; Autoscroll when dragging near the top/bottom of a long document list.
    view := Buffer(16, 0)
    DllCall("GetClientRect", "ptr", LV.Hwnd, "ptr", view.Ptr)
    viewHeight := NumGet(view, 12, "Int")
    if A_TickCount - LastScrollTick > 110 {
        if mouseY < 16
            DllCall("SendMessage", "ptr", LV.Hwnd
                , "uint", 0x115, "uptr", 0, "ptr", 0, "ptr") ; SB_LINEUP
        else if mouseY > viewHeight - 16
            DllCall("SendMessage", "ptr", LV.Hwnd
                , "uint", 0x115, "uptr", 1, "ptr", 0, "ptr") ; SB_LINEDOWN
        LastScrollTick := A_TickCount
    }

    slot := Docs.Length + 1
    for i, doc in Docs {
        if GetRowRect(i, &top, &bottom) {
            if mouseY < (top + bottom) / 2 {
                slot := i
                break
            }
        }
    }
    DragInsert := slot

    if slot <= Docs.Length && GetRowRect(slot, &top, &bottom)
        lineY := top
    else if GetRowRect(Docs.Length, &top, &bottom)
        lineY := bottom
    else
        lineY := 3

    LV.GetPos(&lx, &ly, &lw, &lh)
    LineTargetY := Max(ly + 1, Min(ly + lh - 5, ly + lineY))
    if initial
        LineCurrentY := LineTargetY
    DropLine.Visible := true
    UpdateDropLine()
}

GetRowRect(row, &top, &bottom) {
    global LV, Docs
    if row < 1 || row > Docs.Length
        return false
    rc := Buffer(16, 0)
    NumPut("Int", 0, rc, 0) ; LVIR_BOUNDS
    ok := DllCall("SendMessage", "ptr", LV.Hwnd, "uint", 0x100E
        , "uptr", row - 1, "ptr", rc.Ptr, "ptr") ; LVM_GETITEMRECT
    if !ok
        return false
    top := NumGet(rc, 4, "Int")
    bottom := NumGet(rc, 12, "Int")
    return true
}

UpdateDropLine() {
    global LV, DropLine, LineCurrentY
    LV.GetPos(&lx, &ly, &lw)
    DropLine.Move(lx + 4, Round(LineCurrentY), Max(120, lw - 8), 3)
}

FinishDrag() {
    global Dragging, DragSource, DragInsert, Docs, Anim
    global BlockClicksUntil, HoverRow, DropLine
    if !Dragging
        return
    UpdateDropPosition()
    from := DragSource
    slot := DragInsert
    Dragging := false
    DllCall("ReleaseCapture")
    DropLine.Visible := false
    HoverRow := 0
    BlockClicksUntil := A_TickCount + 350

    if from < 1 || from > Docs.Length
        return

    ; The insertion slot was computed BEFORE removal.
    to := slot > from ? slot - 1 : slot
    to := Max(1, Min(Docs.Length, to))
    if to != from {
        moved := Docs.RemoveAt(from)
        Docs.InsertAt(to, moved)
        DrawDocuments()
        SaveOrder()
        if Anim.Has(moved.id)
            Anim[moved.id].flash := 1.0 ; brief blue landing pulse
    }
    RepaintList()
}

SaveOrder() {
    global Docs, OrderTitles, ConfigFile
    OrderTitles := []
    value := ""
    for doc in Docs {
        OrderTitles.Push(doc.title)
        value .= (value = "" ? "" : "|") doc.title
    }
    try IniWrite(value, ConfigFile, "Tabs", "Order")
}

MouseTick(*) {
    global PaneVisible, Pane, LV, Docs, HoverRow, HoverSince
    global PreviewFor, Dragging
    if !PaneVisible
        return

    if Dragging {
        if !GetKeyState("LButton", "P")
            FinishDrag()
        else
            UpdateDropPosition()
        return
    }

    MouseGetPos(&mx, &my, &mouseWindow, &mouseCtrl, 2)
    row := mouseCtrl = LV.Hwnd ? HitTestRow(mx, my) : 0
    if row != HoverRow {
        HoverRow := row
        HoverSince := A_TickCount
        HidePreview()
    }
    if row > 0 && row <= Docs.Length {
        if A_TickCount - HoverSince > 520
            && PreviewFor != Docs[row].id
            ShowPreview(Docs[row])
    }
}

HitTestRow(mx, my) {
    global LV, Docs
    pt := Buffer(8, 0)
    NumPut("Int", mx, pt, 0)
    NumPut("Int", my, pt, 4)
    DllCall("ScreenToClient", "ptr", LV.Hwnd, "ptr", pt.Ptr)
    hit := Buffer(24, 0) ; LVHITTESTINFO, Win32 24 bytes
    NumPut("Int", NumGet(pt, 0, "Int"), hit, 0)
    NumPut("Int", NumGet(pt, 4, "Int"), hit, 4)
    index := DllCall("SendMessage", "ptr", LV.Hwnd
        , "uint", 0x1012, "uptr", 0, "ptr", hit.Ptr, "ptr")
    return index >= 0 && index < Docs.Length ? index + 1 : 0
}

AnimateTick(*) {
    global Docs, Anim, HoverRow, ActiveId, Dragging
    global LineTargetY, LineCurrentY, PreviewVisible
    global PreviewAnchorX, PreviewCurrentX, PreviewAnchorY, Preview
    global PaneVisible
    if !PaneVisible
        return
    dirty := false

    for i, doc in Docs {
        if !Anim.Has(doc.id)
            continue
        data := Anim[doc.id]
        hTarget := (!Dragging && i = HoverRow) ? 1.0 : 0.0
        aTarget := doc.id = ActiveId ? 1.0 : 0.0
        newHover := EaseTo(data.hover, hTarget, 0.25)
        newActive := EaseTo(data.active, aTarget, 0.23)
        newFlash := data.flash < 0.012 ? 0.0 : data.flash * 0.79
        if Abs(data.hover - newHover) > 0.001
            || Abs(data.active - newActive) > 0.001
            || Abs(data.flash - newFlash) > 0.001
            dirty := true
        data.hover := newHover
        data.active := newActive
        data.flash := newFlash
    }
    if dirty
        RepaintList()

    ; The insertion marker follows with a soft 25ms ease-out animation.
    if Dragging && Abs(LineCurrentY - LineTargetY) > 0.5 {
        LineCurrentY += (LineTargetY - LineCurrentY) * 0.43
        UpdateDropLine()
    }

    ; The original DWM thumbnail smoothly slides into place on hover.
    if PreviewVisible && Abs(PreviewCurrentX - PreviewAnchorX) > 0.5 {
        PreviewCurrentX += (PreviewAnchorX - PreviewCurrentX) * 0.34
        DllCall("SetWindowPos", "ptr", Preview.Hwnd, "ptr", 0
            , "int", Round(PreviewCurrentX), "int", PreviewAnchorY
            , "int", 0, "int", 0, "uint", 0x15)
    }
}

EaseTo(current, target, factor) {
    if Abs(current - target) < 0.012
        return target
    return current + (target - current) * factor
}

MixRGB(c1, c2, amount) {
    amount := Max(0.0, Min(1.0, amount))
    r := Round(((c1 >> 16) & 255) * (1 - amount)
        + ((c2 >> 16) & 255) * amount)
    g := Round(((c1 >> 8) & 255) * (1 - amount)
        + ((c2 >> 8) & 255) * amount)
    b := Round((c1 & 255) * (1 - amount) + (c2 & 255) * amount)
    return (r << 16) | (g << 8) | b
}

ColorRef(rgb) {
    ; Win32 COLORREF is 0x00BBGGRR, unlike the RGB literals above.
    return ((rgb & 255) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 255)
}

RepaintList() {
    global LV
    DllCall("InvalidateRect", "ptr", LV.Hwnd, "ptr", 0, "int", 0)
}

ShowPreview(doc) {
    global Preview, PTitle, PFallback, Pane, ThumbHandle, PreviewFor
    global PreviewVisible, PreviewAnchorX, PreviewCurrentX, PreviewAnchorY
    HidePreview()
    title := doc.title
    PTitle.Text := StrLen(title) > 38 ? SubStr(title, 1, 37) "…" : title

    WinGetPos(&px, &py, &pw, &ph, "ahk_id " Pane.Hwnd)
    w := 320, h := 245
    MonitorGetWorkArea(1, &left, &top, &right, &bottom)
    x := px - w - 12
    if x < left
        x := Min(right - w, px + pw + 12)
    y := Max(top, Min(bottom - h, py + 40))

    PreviewAnchorX := x
    PreviewCurrentX := x + (x < px ? 15 : -15)
    PreviewAnchorY := y

    PFallback.Visible := true
    Preview.Show("NA x" Round(PreviewCurrentX) " y" Round(y)
        " w" w " h" h)
    PreviewVisible := true
    PreviewFor := doc.id

    handle := Buffer(A_PtrSize, 0)
    try {
        hr := DllCall("dwmapi\DwmRegisterThumbnail", "ptr", Preview.Hwnd
            , "ptr", doc.id, "ptr", handle.Ptr, "int")
        if hr != 0
            return
        ThumbHandle := NumGet(handle, 0, "ptr")
        properties := Buffer(48, 0) ; DWM_THUMBNAIL_PROPERTIES
        NumPut("UInt", 0x1 | 0x8, properties, 0)
        NumPut("Int", 10, properties, 4)
        NumPut("Int", 43, properties, 8)
        NumPut("Int", 310, properties, 12)
        NumPut("Int", 230, properties, 16)
        NumPut("Int", 1, properties, 40)
        if DllCall("dwmapi\DwmUpdateThumbnailProperties", "ptr"
                , ThumbHandle, "ptr", properties.Ptr, "int") = 0
            PFallback.Visible := false
    } catch {
        ; Preview failure never affects document switching.
    }
}

HidePreview(*) {
    global PreviewVisible, Preview, ThumbHandle, PreviewFor
    if ThumbHandle {
        try DllCall("dwmapi\DwmUnregisterThumbnail", "ptr", ThumbHandle, "int")
        ThumbHandle := 0
    }
    if PreviewVisible {
        Preview.Hide()
        PreviewVisible := false
    }
    PreviewFor := 0
}

InitialDock() {
    global Pane, ConfigFile, LastWord
    width := Max(205, Min(650
        , Integer(IniRead(ConfigFile, "Pane", "Width", "258"))))
    height := Max(210, Min(1200
        , Integer(IniRead(ConfigFile, "Pane", "Height", "570"))))
    xSaved := IniRead(ConfigFile, "Pane", "X", "")
    ySaved := IniRead(ConfigFile, "Pane", "Y", "")
    if xSaved != "" && ySaved != "" {
        x := Integer(xSaved)
        y := Integer(ySaved)
    } else {
        WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " LastWord)
        x := wx + ww - width - 7
        y := wy + 175
        height := Max(210, wh - 197)
    }
    Pane.Show("NA x" Round(x) " y" Round(y) " w" width " h" height)
}

Redock(*) {
    global Pane, LastWord, PaneVisible
    hwnd := WinActive("ahk_class OpusApp ahk_exe WINWORD.EXE")
    if !hwnd
        hwnd := LastWord
    if !hwnd || !WinExist("ahk_id " hwnd)
        return
    try {
        WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " hwnd)
        WinGetPos(, , &pw, &ph, "ahk_id " Pane.Hwnd)
        if pw < 205
            pw := 258
        Pane.Show("NA x" Round(wx + ww - pw - 7)
            " y" Round(wy + 175) " w" pw " h" ph)
        PaneVisible := true
        SavePosition()
    } catch {
    }
}

Toggle(*) {
    global Enabled, Pane, PaneVisible
    Enabled := !Enabled
    if !Enabled {
        SavePosition()
        HidePreview()
        Pane.Hide()
        PaneVisible := false
    } else
        Refresh()
}

SavePosition(*) {
    global PaneVisible, Pane, ConfigFile, LastSaved
    if !PaneVisible
        return
    try {
        WinGetPos(&x, &y, &w, &h, "ahk_id " Pane.Hwnd)
        key := x "," y "," w "," h
        if key = LastSaved
            return
        IniWrite(x, ConfigFile, "Pane", "X")
        IniWrite(y, ConfigFile, "Pane", "Y")
        IniWrite(w, ConfigFile, "Pane", "Width")
        IniWrite(h, ConfigFile, "Pane", "Height")
        LastSaved := key
    } catch {
    }
}

OnQuit(*) {
    global Dragging
    if Dragging
        DllCall("ReleaseCapture")
    SavePosition()
    HidePreview()
}

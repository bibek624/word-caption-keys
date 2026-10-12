#Requires AutoHotkey v2.0
#SingleInstance Force

; Word Caption Keys - fast captions and cross-references in Microsoft Word
;   Alt+F        list figures, Enter inserts a cross-reference (label and number)
;   Alt+T        list tables,  Enter inserts a cross-reference (label and number)
;   Ctrl+Alt+F   insert a new figure caption at the cursor
;   Ctrl+Alt+T   insert a new table caption
; The keys only act while a Word window is active.

A_IconTip := "Word Caption Keys"
global WCK_LINK := A_Startup "\Word Caption Keys.lnk"
A_TrayMenu.Delete()
A_TrayMenu.Add("Shortcuts", (*) => MsgBox(
    "In Word:`n`n"
    "Alt+F`t`tCross-reference a figure`n"
    "Alt+T`t`tCross-reference a table`n"
    "Ctrl+Alt+F`tNew figure caption`n"
    "Ctrl+Alt+T`tNew table caption", "Word Caption Keys"))
A_TrayMenu.Add("Start with Windows", WCK_ToggleStartup)
A_TrayMenu.Add()
A_TrayMenu.Add("Exit", (*) => ExitApp())
if FileExist(WCK_LINK)
    A_TrayMenu.Check("Start with Windows")

WCK_ToggleStartup(*) {
    if FileExist(WCK_LINK) {
        FileDelete(WCK_LINK)
        A_TrayMenu.Uncheck("Start with Windows")
    } else {
        if A_IsCompiled
            FileCreateShortcut(A_ScriptFullPath, WCK_LINK)
        else
            FileCreateShortcut(A_AhkPath, WCK_LINK, , '"' A_ScriptFullPath '"')
        A_TrayMenu.Check("Start with Windows")
    }
}
; ================= Word: cross-references and captions =================
; Alt+F / Alt+T list the document's figures / tables and insert a cross-reference (label and number).
; Ctrl+Alt+F / Ctrl+Alt+T insert a new figure / table caption at the cursor.
global XR_Gui := "", XR_Edit := "", XR_LV := ""
global XR_All := [], XR_Shown := [], XR_Label := "", XR_Hwnd := 0

OnMessage(0x6, XR_OnActivate)         ; hide list when it loses focus

#HotIf WinActive("ahk_class OpusApp")
!f:: XR_Show("Figure")                ; Alt+F, cross-reference a figure
!t:: XR_Show("Table")                 ; Alt+T, cross-reference a table
^!f:: WD_Caption("Figure")            ; Ctrl+Alt+F, new figure caption
^!t:: WD_Caption("Table")             ; Ctrl+Alt+T, new table caption
#HotIf

#HotIf XR_GuiActive()
Up:: XR_Move(-1)
Down:: XR_Move(1)
Enter:: XR_Pick()
NumpadEnter:: XR_Pick()
#HotIf

XR_GuiActive() {
    return IsObject(XR_Gui) && WinActive("ahk_id " XR_Gui.Hwnd)
}

WD_App() {
    try return ComObjActive("Word.Application")
    return ""
}

WD_Tip(msg) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -2000)
}

WD_Caption(label) {
    word := WD_App()
    if !IsObject(word)
        return WD_Tip("Could not reach Word")
    try WD_InsertCaption(word, label)
    catch
        WD_Tip("Could not insert the caption (is the document editable?)")
}

; Insert "<label> <n>" as a real Word caption and leave the cursor right after the number
WD_InsertCaption(word, label) {
    sel := word.Selection
    pos := label = "Table" ? 0 : 1            ; tables are captioned above, figures below
    if label = "Table" && sel.Information[12] {       ; cursor inside a table: caption that table
        sel.Tables.Item(1).Select()
    } else if sel.Type = 1 || sel.Type = 2 {          ; plain cursor or selected text (not a picture)
        rng := sel.Paragraphs.Item(1).Range
        if rng.End - rng.Start > 1 {                  ; paragraph has text: start a new one below it
            rng.InsertParagraphAfter()
            sel.SetRange(rng.End - 1, rng.End - 1)
        } else {
            sel.Collapse(1)
        }
    }
    sel.InsertCaption(label, "", "", pos)
}

XR_Show(label) {
    global XR_Gui, XR_Edit, XR_LV, XR_All, XR_Label, XR_Hwnd
    XR_Hwnd := WinExist("A")
    XR_Label := label
    XR_All := []
    word := WD_App()
    if !IsObject(word)
        return WD_Tip("Could not reach Word")
    try {
        for text in word.ActiveDocument.GetCrossReferenceItems(label)
            XR_All.Push({idx: A_Index, text: Trim(text, " `t`r`n")})
    }
    if !XR_All.Length
        return WD_Tip("No " StrLower(label) " captions in this document")

    if !IsObject(XR_Gui) {
        XR_Gui := Gui("+AlwaysOnTop -Caption +Border +ToolWindow", "Cross-reference")
        XR_Gui.BackColor := "1E1E1E"
        XR_Gui.MarginX := 12, XR_Gui.MarginY := 12
        XR_Gui.SetFont("s13 cWhite", "Segoe UI")
        XR_Edit := XR_Gui.AddEdit("w620 Background2D2D2D -E0x200")
        XR_Edit.OnEvent("Change", (*) => XR_Filter())
        XR_Gui.SetFont("s10 cE0E0E0", "Segoe UI")
        XR_LV := XR_Gui.AddListView("w620 r14 -Multi -Hdr Background252525", ["Caption"])
        XR_LV.ModifyCol(1, 598)
        XR_LV.OnEvent("DoubleClick", (*) => XR_Pick())
        XR_Gui.OnEvent("Escape", (*) => XR_Gui.Hide())
    }

    XR_Edit.Value := ""
    XR_Filter()
    GuiShowOnMonitorOf(XR_Gui, XR_Hwnd, 644)
    XR_Edit.Focus()
}

XR_Filter() {
    global XR_Shown
    q := Trim(XR_Edit.Value)
    terms := StrSplit(q, " ", " ")
    exact := [], rest := []
    for item in XR_All {
        ok := true
        for t in terms
            if t != "" && !InStr(item.text, t) {
                ok := false
                break
            }
        if !ok
            continue
        ; typing just the number puts that exact caption first ("1" -> Figure 1 before Figure 10)
        if q != "" && RegExMatch(item.text, "i)^\S+\s+\Q" q "\E(?![\w-])")
            exact.Push(item)
        else
            rest.Push(item)
    }
    XR_Shown := exact
    for item in rest
        XR_Shown.Push(item)

    XR_LV.Opt("-Redraw")
    XR_LV.Delete()
    for item in XR_Shown
        XR_LV.Add(, item.text)
    XR_LV.Opt("+Redraw")
    if XR_Shown.Length
        XR_LV.Modify(1, "Select Focus")
}

XR_Move(d) {
    n := XR_LV.GetCount()
    if !n
        return
    r := Max(1, Min(n, XR_LV.GetNext() + d))
    XR_LV.Modify(0, "-Select -Focus")
    XR_LV.Modify(r, "Select Focus Vis")
}

XR_Pick() {
    row := XR_LV.GetNext()
    if !row || row > XR_Shown.Length
        return
    idx := XR_Shown[row].idx
    XR_Gui.Hide()
    if !WinExist(XR_Hwnd)
        return
    WinActivate(XR_Hwnd)
    WinWaitActive(XR_Hwnd, , 1)
    try WD_App().Selection.InsertCrossReference(XR_Label, 3, idx, true, false)   ; 3 = only label and number
    catch
        WD_Tip("Could not insert the cross-reference")
}

XR_OnActivate(wParam, lParam, msg, hwnd) {
    if IsObject(XR_Gui) && hwnd = XR_Gui.Hwnd && (wParam & 0xFFFF) = 0
        XR_Gui.Hide()
}

; Show a GUI centred on the monitor that holds the given window
GuiShowOnMonitorOf(g, hwnd, width) {
    mon := MonitorGetPrimary()
    try {
        WinGetPos(&ax, &ay, &aw, &ah, hwnd)
        cx := ax + aw // 2, cy := ay + ah // 2
        Loop MonitorGetCount() {
            MonitorGet(A_Index, &l, &t, &r, &b)
            if cx >= l && cx < r && cy >= t && cy < b {
                mon := A_Index
                break
            }
        }
    }
    MonitorGetWorkArea(mon, &l, &t, &r, &b)
    g.Show("Hide w" width)            ; size it first so it can be centred
    g.GetPos(, , &gw, &gh)
    g.Show(Format("x{} y{}", l + (r - l - gw) // 2, t + (b - t - gh) // 3))
}

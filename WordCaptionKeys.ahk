#Requires AutoHotkey v2.0
#SingleInstance Force

; Word Caption Keys - fast captions and cross-references in Microsoft Word
;
; Default keys (change them in WordCaptionKeys.ini, created next to this script on first run):
;   Alt+F        list figures,  Enter inserts a cross-reference (label and number)
;   Alt+T        list tables,   Enter inserts a cross-reference (label and number)
;   Alt+H        list headings, Enter inserts a cross-reference (heading number, or its text if unnumbered)
;   Ctrl+Alt+F   insert a new figure caption at the cursor
;   Ctrl+Alt+T   insert a new table caption
;   Ctrl+Alt+U   update every field in the document
; The keys only act while a Word window is active.

global WCK_INI := A_ScriptDir "\WordCaptionKeys.ini"
global WCK_ACTIONS := [               ; settings name, default key, description, action
    ["FigureReference",  "!f",  "Cross-reference a figure",  (*) => XR_Show("figure")],
    ["TableReference",   "!t",  "Cross-reference a table",   (*) => XR_Show("table")],
    ["HeadingReference", "!h",  "Cross-reference a heading", (*) => XR_Show("heading")],
    ["FigureCaption",    "^!f", "New figure caption",        (*) => WD_Caption("figure")],
    ["TableCaption",     "^!t", "New table caption",         (*) => WD_Caption("table")],
    ["UpdateAllFields",  "^!u", "Update all fields",         (*) => WD_UpdateAll()]
]
global WCK_Keys := Map(), WCK_FigureLabels := [], WCK_TableLabels := []
global WCK_LINK := A_Startup "\Word Caption Keys.lnk"

global XR_Gui := "", XR_Edit := "", XR_LV := ""
global XR_All := [], XR_Shown := [], XR_Kind := "", XR_Hwnd := 0

global WCK_FirstRun := !FileExist(WCK_INI)
WCK_LoadSettings()
WCK_BuildTray()
if WCK_FirstRun
    WCK_Welcome()
OnMessage(0x6, XR_OnActivate)         ; hide list when it loses focus

#HotIf XR_GuiActive()
Up:: XR_Move(-1)
Down:: XR_Move(1)
Enter:: XR_Pick(false)
NumpadEnter:: XR_Pick(false)
^Enter:: XR_Pick(true)                ; headings: insert the heading text instead of its number
^NumpadEnter:: XR_Pick(true)
#HotIf

; ================= Settings and tray =================
WCK_LoadSettings() {
    global WCK_FigureLabels, WCK_TableLabels
    if !FileExist(WCK_INI)
        FileAppend(WCK_DefaultIni(), WCK_INI, "UTF-16")

    bad := ""
    HotIfWinActive("ahk_class OpusApp")
    for a in WCK_ACTIONS {
        key := Trim(IniRead(WCK_INI, "Keys", a[1], a[2]))
        WCK_Keys[a[1]] := key
        if key = ""                   ; blank = shortcut switched off
            continue
        try Hotkey(key, a[4])
        catch {
            bad .= "`n" a[1] " = " key
            WCK_Keys[a[1]] := ""
        }
    }
    HotIfWinActive()
    WCK_FigureLabels := StrSplit(IniRead(WCK_INI, "Labels", "FigureLabels", "Fig"), ",", " `t")
    WCK_TableLabels := StrSplit(IniRead(WCK_INI, "Labels", "TableLabels", "Tab"), ",", " `t")
    if bad != ""
        MsgBox("These shortcuts in WordCaptionKeys.ini are not valid keys and were skipped:`n" bad, "Word Caption Keys")
}

WCK_DefaultIni() {
    text := "
    (
; Word Caption Keys settings. After changing anything here,
; right-click the tray icon and choose Reload.
;
; Key symbols:  ! = Alt   ^ = Ctrl   + = Shift   # = Windows key
; Examples:     !f = Alt+F     ^!f = Ctrl+Alt+F     ^+r = Ctrl+Shift+R     F8 = the F8 key
; Leave a line empty after the = sign to switch that shortcut off.

[Keys]
FigureReference=!f
TableReference=!t
HeadingReference=!h
FigureCaption=^!f
TableCaption=^!t
UpdateAllFields=^!u

[Labels]
; Word's own Figure and Table labels are always used, in any language.
; Custom caption labels whose names start with one of these (comma separated)
; are treated as figures or tables too, for example "Fig." or "Tab.".
FigureLabels=Fig
TableLabels=Tab

    )"
    return StrReplace(text, "`n", "`r`n")
}

WCK_BuildTray() {
    A_IconTip := "Word Caption Keys"
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Shortcuts", WCK_ShowKeys)
    A_TrayMenu.Add("Settings", (*) => Run('notepad.exe "' WCK_INI '"'))
    A_TrayMenu.Add("Reload", (*) => Reload())
    A_TrayMenu.Add("Start with Windows", WCK_ToggleStartup)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    A_TrayMenu.Default := "Shortcuts"
    if FileExist(WCK_LINK)
        A_TrayMenu.Check("Start with Windows")
}

WCK_KeysText() {
    text := "In Word:`n"
    for a in WCK_ACTIONS
        text .= "`n" (WCK_Keys[a[1]] = "" ? "(off)" : WCK_Pretty(WCK_Keys[a[1]])) "`t" a[3]
    text .= "`n`nIn a list: type to filter, Enter to insert, Esc to cancel."
    text .= "`nHeadings: Ctrl+Enter inserts the heading text instead of its number."
    return text
}

WCK_ShowKeys(*) {
    MsgBox(WCK_KeysText(), "Word Caption Keys")
}

; Shown once, the first time the script runs
WCK_Welcome() {
    text := "Word Caption Keys is now running. Look for the green H icon near the clock.`n`n"
        . WCK_KeysText()
        . "`n`nStart Word Caption Keys automatically every time Windows starts?"
    if MsgBox(text, "Word Caption Keys", "YesNo Iconi") = "Yes" && !FileExist(WCK_LINK)
        WCK_ToggleStartup()
}

; "^!f" -> "Ctrl+Alt+F"
WCK_Pretty(key) {
    static names := Map("^", "Ctrl+", "!", "Alt+", "+", "Shift+", "#", "Win+")
    mods := ""
    while StrLen(key) > 1 && names.Has(SubStr(key, 1, 1)) {
        mods .= names[SubStr(key, 1, 1)]
        key := SubStr(key, 2)
    }
    return mods (StrLen(key) = 1 ? StrUpper(key) : key)
}

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

; ================= Word helpers =================
WD_App() {
    try return ComObjActive("Word.Application")
    return ""
}

WD_Tip(msg) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -2000)
}

; Caption labels that count as figures or tables: Word's built-in one (whatever it is
; called in this language) plus custom labels matching the prefixes in the settings file
WD_Labels(word, kind) {
    names := []
    try names.Push(word.CaptionLabels.Item(kind = "table" ? -2 : -1).Name)
    try {
        labels := word.CaptionLabels
        Loop labels.Count {
            lab := labels.Item(A_Index)
            if lab.BuiltIn
                continue
            for prefix in (kind = "table" ? WCK_TableLabels : WCK_FigureLabels)
                if prefix != "" && InStr(lab.Name, prefix) = 1 {
                    names.Push(lab.Name)
                    break
                }
        }
    }
    if !names.Length
        names.Push(kind = "table" ? "Table" : "Figure")
    return names
}

; Captions (or headings, for ref = 1) that Word offers as cross-reference targets
WD_Items(doc, ref) {
    items := []
    try {
        for text in doc.GetCrossReferenceItems(ref)
            items.Push(text)
    }
    return items
}

WD_Caption(kind) {
    word := WD_App()
    if !IsObject(word)
        return WD_Tip("Could not reach Word")
    try {
        ; use the label this document already uses most, so numbering stays in one sequence
        label := "", most := -1
        for name in WD_Labels(word, kind) {
            n := WD_Items(word.ActiveDocument, name).Length
            if n > most
                label := name, most := n
        }
        WD_InsertCaption(word, label, kind = "table")
    } catch {
        WD_Tip("Could not insert the caption (is the document editable?)")
    }
}

; Insert "<label> <n>" as a real Word caption and leave the cursor right after the number
WD_InsertCaption(word, label, isTable) {
    sel := word.Selection
    if isTable && sel.Information[12] {               ; cursor inside a table: caption that table
        sel.Tables.Item(1).Select()
    } else if sel.Type = 1 || sel.Type = 2 {          ; plain cursor or selected text (not a picture)
        rng := sel.Paragraphs.Item(1).Range
        if rng.End - rng.Start > 1 {                  ; paragraph has text: start a new one below it,
            sel.SetRange(rng.End - 1, rng.End - 1)    ; like pressing Enter at the end of the line
            sel.TypeParagraph()
        } else {
            sel.Collapse(1)
        }
    }
    sel.InsertCaption(label, "", "", isTable ? 0 : 1) ; tables are captioned above, figures below
}

; Refresh cross-references, caption numbers, tables of contents, page numbers: every field,
; in the body, headers, footers, footnotes and text boxes
WD_UpdateAll() {
    word := WD_App()
    if !IsObject(word)
        return WD_Tip("Could not reach Word")
    try {
        doc := word.ActiveDocument
        word.ScreenUpdating := false
        try {
            stories := doc.StoryRanges
            Loop 17 {                                 ; story types: main text, footnotes, headers, ...
                try {
                    rng := stories.Item(A_Index)
                    Loop 500 {
                        rng.Fields.Update()
                        rng := rng.NextStoryRange
                        if !IsObject(rng) || !ComObjValue(rng)
                            break
                    }
                }
            }
        } finally {
            word.ScreenUpdating := true
        }
        WD_Tip("All fields updated")
    } catch {
        WD_Tip("Could not update the fields (is the document editable?)")
    }
}

; ================= Cross-reference picker =================
XR_GuiActive() {
    return IsObject(XR_Gui) && WinActive("ahk_id " XR_Gui.Hwnd)
}

XR_Show(kind) {
    global XR_Gui, XR_Edit, XR_LV, XR_All, XR_Kind, XR_Hwnd
    XR_Hwnd := WinExist("A")
    XR_Kind := kind
    XR_All := []
    word := WD_App()
    if !IsObject(word)
        return WD_Tip("Could not reach Word")
    try {
        doc := word.ActiveDocument
        if kind = "heading" {
            for text in WD_Items(doc, 1)              ; leading spaces show the heading level, keep them
                XR_All.Push({ref: 1, idx: A_Index, text: RTrim(text, " `t`r`n")})
        } else {
            for label in WD_Labels(word, kind)
                for text in WD_Items(doc, label)
                    XR_All.Push({ref: label, idx: A_Index, text: Trim(text, " `t`r`n")})
        }
    }
    if !XR_All.Length
        return WD_Tip(kind = "heading" ? "No headings in this document" : "No " kind " captions in this document")

    if !IsObject(XR_Gui) {
        XR_Gui := Gui("+AlwaysOnTop -Caption +Border +ToolWindow", "Cross-reference")
        XR_Gui.BackColor := "1E1E1E"
        XR_Gui.MarginX := 12, XR_Gui.MarginY := 12
        XR_Gui.SetFont("s13 cWhite", "Segoe UI")
        XR_Edit := XR_Gui.AddEdit("w620 Background2D2D2D -E0x200")
        XR_Edit.OnEvent("Change", (*) => XR_Filter())
        XR_Gui.SetFont("s10 cE0E0E0", "Segoe UI")
        XR_LV := XR_Gui.AddListView("w620 r14 -Multi -Hdr Background252525", ["Item"])
        XR_LV.ModifyCol(1, 598)
        XR_LV.OnEvent("DoubleClick", (*) => XR_Pick(false))
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
    ; typing just the number puts that exact item first ("1" -> Figure 1 before Figure 10)
    exactRe := "i)^\s*" (XR_Kind = "heading" ? "" : "\S+\s+") "\Q" q "\E(?![\w.-])"
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
        if q != "" && RegExMatch(item.text, exactRe)
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

XR_Pick(textOnly) {
    row := XR_LV.GetNext()
    if !row || row > XR_Shown.Length
        return
    item := XR_Shown[row]
    XR_Gui.Hide()
    if !WinExist(XR_Hwnd)
        return
    WinActivate(XR_Hwnd)
    WinWaitActive(XR_Hwnd, , 1)
    try XR_Insert(WD_App().Selection, XR_Kind, item, textOnly)
    catch
        WD_Tip("Could not insert the cross-reference")
}

XR_Insert(sel, kind, item, textOnly) {
    if kind != "heading" {
        sel.InsertCrossReference(item.ref, 3, item.idx, true, false)      ; 3 = only label and number
    } else if textOnly {
        sel.InsertCrossReference(1, -1, item.idx, true, false)            ; -1 = heading text
    } else {
        sel.InsertCrossReference(1, -4, item.idx, true, false)            ; -4 = heading number
        sel.MoveLeft(1, 1, 1)                                             ; select the new field
        if sel.Fields.Count = 1 && (sel.Text = "0" || sel.Text = "") {    ; heading is not numbered
            sel.Delete()
            sel.InsertCrossReference(1, -1, item.idx, true, false)
        } else {
            sel.Collapse(0)
        }
    }
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

; Press F1
F1::
    Selection := Explorer_GetSelection()
    If (Selection) && (WhereTo := SelectFolderEx()) {
        For Each, Item in StrSplit(Selection, "|") {
            If !InStr(FileExist(Item), "D") {
                FileMove, % Item, % WhereTo
            }
        }
    }
Return

; https://www.autohotkey.com/boards/viewtopic.php?p=255256#p255256
Explorer_GetSelection() {
    WinGetClass, winClass, % "ahk_id" . hWnd := WinExist("A")
    if !(winClass ~="Progman|WorkerW|(Cabinet|Explore)WClass")
        Return
    shellWindows := ComObjCreate("Shell.Application").Windows
    if (winClass ~= "Progman|WorkerW")
        shellFolderView := shellWindows.FindWindowSW(0, 0, SWC_DESKTOP := 8, 0, SWFO_NEEDDISPATCH := 1).Document
    else {
        for window in shellWindows
            if (hWnd = window.HWND) && (shellFolderView := window.Document)
            break
    }
    for item in shellFolderView.SelectedItems
        result .= (result = "" ? "" : "|") . item.Path
    Return result
}

; https://www.autohotkey.com/boards/viewtopic.php?t=18939
SelectFolderEx(StartingFolder := "", Prompt := "", OwnerHwnd := 0, OkBtnLabel := "") {
    Static OsVersion := DllCall("GetVersion", "UChar")
         , IID_IShellItem := 0
         , InitIID := VarSetCapacity(IID_IShellItem, 16, 0)
         & DllCall("Ole32.dll\IIDFromString", "WStr", "{43826d1e-e718-42ee-bc55-a1e261c37bfe}", "Ptr", &IID_IShellItem)
         , Show := A_PtrSize * 3
         , SetOptions := A_PtrSize * 9
         , SetFolder := A_PtrSize * 12
         , SetTitle := A_PtrSize * 17
         , SetOkButtonLabel := A_PtrSize * 18
         , GetResult := A_PtrSize * 20
    SelectedFolder := ""
    If (OsVersion < 6) { ; IFileDialog requires Win Vista+, so revert to FileSelectFolder
        FileSelectFolder, SelectedFolder, *%StartingFolder%, 3, %Prompt%
        Return SelectedFolder
    }
    OwnerHwnd := DllCall("IsWindow", "Ptr", OwnerHwnd, "UInt") ? OwnerHwnd : 0
    If !(FileDialog := ComObjCreate("{DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7}", "{42f85136-db7e-439c-85f1-e4075d135fc8}"))
        Return ""
    VTBL := NumGet(FileDialog + 0, "UPtr")
    ; FOS_CREATEPROMPT | FOS_NOCHANGEDIR | FOS_PICKFOLDERS
    DllCall(NumGet(VTBL + SetOptions, "UPtr"), "Ptr", FileDialog, "UInt", 0x00002028, "UInt")
    If (StartingFolder <> "")
        If !DllCall("Shell32.dll\SHCreateItemFromParsingName", "WStr", StartingFolder, "Ptr", 0, "Ptr", &IID_IShellItem, "PtrP", FolderItem)
        DllCall(NumGet(VTBL + SetFolder, "UPtr"), "Ptr", FileDialog, "Ptr", FolderItem, "UInt")
    If (Prompt <> "")
        DllCall(NumGet(VTBL + SetTitle, "UPtr"), "Ptr", FileDialog, "WStr", Prompt, "UInt")
    If (OkBtnLabel <> "")
        DllCall(NumGet(VTBL + SetOkButtonLabel, "UPtr"), "Ptr", FileDialog, "WStr", OkBtnLabel, "UInt")
    If !DllCall(NumGet(VTBL + Show, "UPtr"), "Ptr", FileDialog, "Ptr", OwnerHwnd, "UInt") {
        If !DllCall(NumGet(VTBL + GetResult, "UPtr"), "Ptr", FileDialog, "PtrP", ShellItem, "UInt") {
            GetDisplayName := NumGet(NumGet(ShellItem + 0, "UPtr"), A_PtrSize * 5, "UPtr")
            If !DllCall(GetDisplayName, "Ptr", ShellItem, "UInt", 0x80028000, "PtrP", StrPtr) ; SIGDN_DESKTOPABSOLUTEPARSING
                SelectedFolder := StrGet(StrPtr, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "Ptr", StrPtr)
            ObjRelease(ShellItem)
        } 
    }
    If (FolderItem)
        ObjRelease(FolderItem)
    ObjRelease(FileDialog)
    Return SelectedFolder
}

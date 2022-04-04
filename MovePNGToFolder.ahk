; Press F1
F1::
    If (Selection := Explorer_GetSelection()) {
        For Each, FilePath in StrSplit(Selection, "|") {
            If !InStr(FileExist(FilePath), "D") {
                SplitPath, % FilePath,,,, OutNameNoExt
                Folder := SubStr(OutNameNoExt, InStr(OutNameNoExt, " ",, 0) + 1)
                If !FileExist(Folder) {
                    FileCreateDir, % Folder
                } 
                FileMove, % FilePath, % Folder
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

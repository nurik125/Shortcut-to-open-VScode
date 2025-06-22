GetActiveExplorerPath() {
        hwnd := WinExist("A")
        for window in ComObject("Shell.Application").Windows {
            try {
                if (window.hwnd = hwnd)
                    return window.Document.Folder.Self.Path
            }
        }
        return false
    }

!c::{
    dirPath := GetActiveExplorerPath()
    if dirPath{
        Run Format('cmd /c code "{}"', GetActiveExplorerPath())
    }
}
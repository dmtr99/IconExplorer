#SingleInstance force
#Requires AutoHotKey v2.0-

filepath := "C:\Windows\System32\shell32.dll"
KeyWordFile := A_ScriptDir "\KeyWords.ini"
MyGui := Gui()    ; Create a MyGui window.
ogEdit_Search := MyGui.AddEdit("w300")
LV := MyGui.AddListView("h500 w600", ["Icon & Number", "File","KeyWords"])    ; Create a ListView.
LV.Opt("-Multi")
ogPic_Icon64 := MyGui.AddPicture("w64 h-1 icon1 section", "shell32.dll")
ogPic_Icon40 := MyGui.AddPicture("x+10 yp w40 h-1 icon1", "shell32.dll")
ogPic_Icon32 := MyGui.AddPicture("x+10 yp w32 h-1 icon1", "shell32.dll")
ogPic_Icon24 := MyGui.AddPicture("x+10 yp w24 h-1 icon1", "shell32.dll")
ogPic_Icon16 := MyGui.AddPicture("x+10 yp w16 h-1 icon1", "shell32.dll")
ogEdit_KeyWords := MyGui.AddEdit("xm ys+74 w600")
ogEdit_KeyWords.OnEvent("Change", ButtonSave)

ogEdit_Search.OnEvent("Change", (*)=> LV_Update(LV))

if(false){
    LV.Opt("+Icon")
    ImageListID := IL_Create(1,200,1)    ; Create an ImageList to hold 10 small icons.
} else{
    ImageListID := IL_Create(1,200,0)    ; Create an ImageList to hold 10 small icons.
}

LV.SetImageList(ImageListID)    ; Assign the above ImageList to the current ListView.
LV.OnEvent("Click", LV_Click)
LV.OnEvent("ContextMenu", LV_ContextMenu)
LV.ModifyCol()    ; Auto-adjust the column widths.
LV_Update(LV)
MyGui.Show

LV_Click(LV, RowNumber){
    Number := StrReplace(LV.GetText(RowNumber),"#","")
    IconFile := LV.GetText(RowNumber,2)
    keyWord := IniRead(KeyWordFile, IconFile, Number, "_")
    ogEdit_KeyWords.Text := keyWord
    ogPic_Icon64.Value := "*icon" Number " " IconFile
    ogPic_Icon40.Value := "*icon" Number " " IconFile
    ogPic_Icon32.Value := "*icon" Number " " IconFile
    ogPic_Icon24.Value := "*icon" Number " " IconFile
    ogPic_Icon16.Value := "*icon" Number " " IconFile
    ; ToolTip(IconFile " [" Number "]")
}

LV_Update(LV){
    aIconFiles := ["shell32.dll", "wmploc.dll", "pifmgr.dll","compstui.dll","ddores.dll","ieframe.dll", "imageres.dll","mmcndmgr.dll","moricons.dll","netshell.dll","pnidui.dll","wpdshext.dll","comres.dll","dmdskres.dll","dsuiext.dll","inetcpl.cpl","mstsc.exe","mstscax.dll","setupapi.dll","shdocvw.dll","urlmon.dll","wiashext.dll","mmres.dll"]
    ; aIconFiles := ["shell32.dll", "wmploc.dll", "pifmgr.dll"]

    LV.Opt("-Redraw")

    Loop {
        LV.Delete()
        SearchValue := ogEdit_Search.Value
        for IconFile in aIconFiles
        {
            Loop 400 {
                if (SearchValue != ogEdit_Search.Value){
                    break
                }
                KeyWord := IniRead(KeyWordFile, IconFile, A_Index, "_")
                if (ogEdit_Search.Text!="" and !InStr(KeyWord,SearchValue)){
                    continue
                }
                hBitmap := LoadPicture(IconFile, "Icon" A_Index)
                if (hBitmap = 0) {
                    break
                }
                IconIndex := IL_Add(ImageListID, "HBITMAP:" hBitmap)
                LV.Add("Icon" . IconIndex, "#" A_Index, IconFile, KeyWord)
            }
        }
        if (SearchValue = ogEdit_Search.Value){
            break
        }
    }
    LV.ModifyCol()
    LV.Opt("+Redraw")
}

ButtonSave(*){
    RowNumber := LV.GetNext(0, "F")
    Number := StrReplace(LV.GetText(RowNumber), "#", "")
    File := LV.GetText(RowNumber, 2)
    IniWrite(ogEdit_KeyWords.Text, KeyWordFile, File, Number)
    LV.Modify(RowNumber, "", "#" Number, File , ogEdit_KeyWords.Text)
}

LV_ContextMenu(LV, RowNumber, IsRightClick, X, Y){

    Number := StrReplace(LV.GetText(RowNumber), "#", "")
    File := LV.GetText(RowNumber, 2)

    MyMenu := Menu()
    Text := '"' File '", ' Number
    MyMenu.add( "Copy [" Text "]", (*) => (A_Clipboard := Text, Tooltip2("Copied [" A_Clipboard "]")))
    MyMenu.Show
}

Tooltip2(Text := "", X := "", Y := "", WhichToolTip := "1") {
    ; ToolTip(Text, X, Y, WhichToolTip)
    ToolTip(Text)
    SetTimer () => ToolTip(), -3000
}

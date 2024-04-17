#Requires AutoHotkey v2.0

CSV_Load(FileName, Delimiter := "`,")
{
    if (!FileExist(FileName)) {
        MsgBox("File Not Found")
        return -1
    }

    global dataObj := {
        CSV_TotalRows: 0,
        CSV_TotalCols: 0,
        CSV_Delimiter: "",
        CSV_Path: "",
        CSV_FileName: "",
        CSV_FileNamePath: "",
        dataArr: []
    }

    Local Row
    Local Col

    fileContents := FileRead(FileName)

    fileContents := StrReplace(fileContents, "`r`n`r`n, `r`n")  ;Removes all blank lines from the text in a variable.

    ; NewlineCheck := SubStr(fileContents, -1)
    ; if (NewlineCheck == "`n") {
    ;     fileContents := SubStr(fileContents, -1)
    ; }

    Loop parse, fileContents, "`n", "`r"  ; 在 `r 之前指定 `n, 这样可以同时支持对 Windows 和 Unix 文件的解析.
    {
        If (A_LoopField = "") ; added to skip empty lines
            Continue            ; added

        ; MsgBox("Row: " A_Index ", Data: " A_LoopField)

        local rowArray := []
        dataObj.dataArr.Push(rowArray)
        Col := ReturnDSVArray(A_LoopField, rowArray, Delimiter)
        Row := A_Index

    }

    dataObj.CSV_TotalRows := Row
    dataObj.CSV_TotalCols := Col
    dataObj.CSV_Delimiter := Delimiter
    SplitPath FileName dataObj.CSV_FileName dataObj.CSV_Path
    ; IfNotInString, FileName, `\
    if (!InStr(FileName, "\"))
    {
        dataObj.CSV_FileName := FileName
        dataObj.CSV_Path := A_ScriptDir
    }
    dataObj.CSV_FileNamePath := dataObj.CSV_Path . "\" . dataObj.CSV_FileName
}

CSV_ReadCell(Row, Col)
{
    return dataObj.dataArr[Row][Col]
}

CSV_TotalRows() {
    return dataObj.CSV_TotalRows
}

CSV_TotalCols() {
    return dataObj.CSV_TotalCols
}

ReturnDSVArray(CurrentDSVLine, RowArray := [], Delimiter := ",", Encapsulator := '`"')
{
    global
    if ((StrLen(Delimiter) != 1) || (StrLen(Encapsulator) != 1))
    {
        return -1                            ; return -1 indicating an error ...
    }

    local d := "x" Format("{:X}", Ord(delimiter))       ; used as hex notation in the RegExNeedle
    local e := "x" Format("{:X}", Ord(Encapsulator))    ; used as hex notation in the RegExNeedle

    local p0 := 1                            ; Start of search at char p0 in DSV Line
    local fieldCount := 0                    ; start off with empty fields.
    CurrentDSVLine .= delimiter              ; Add delimiter, otherwise last field won't get recognized

    Loop
    {
        Local RegExNeedle := "\" d "(?=(?:[^\" e "]*\" e "[^\" e "]*\" e ")*(?![^\" e "]*\" e "))"
        Local p1 := RegExMatch(CurrentDSVLine, RegExNeedle, &tmp, p0)
        ; p1 contains now the position of our current delimiter in a 1-based index
        fieldCount++                         ; add count
        local field := SubStr(CurrentDSVLine, p0, p1 - p0)
        ; This is the Line you'll have to change if you want different treatment
        ; otherwise your resulting fields from the DSV data Line will be stored in AHK array
        if (SubStr(field, 1, 1) = Encapsulator)
        {
            ; This is the exception handling for removing any doubled encapsulators and
            ; leading/trailing encapsulator chars
            field := RegExReplace(field, "^\" e "|\" e "$")
            ; StringReplace, field, field, %Encapsulator Encapsulator, %Encapsulator%, All
            field := StrReplace(field, Encapsulator "" Encapsulator, Encapsulator)
        }
        ; Local _field := ReturnArray A_Index  ; construct a reference for our ReturnArray name
        ; %_field% := field                    ; dereference _field and assign our value to it

        RowArray.Push(field)  ;Push the Col data

        ; MsgBox("Row: "  dataObj.dataArr.Length ", Col: " A_Index ", Data: " field)

        if (p1 = 0)
        {                                    ; p1 is 0 when no more delimiter chars have been found
            fieldCount--                     ; so correct fieldCount due to last appended delimiter
            Break                            ; and exit loop
        } Else
            p0 := p1 + 1                     ; set the start of our RegEx Search to last result
    }                                        ; added by one
    return fieldCount
}
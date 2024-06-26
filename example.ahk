﻿#Requires AutoHotkey v2.0
#SingleInstance force

#IncludeAgain csv.ahk

; A basic example illustrating some functions to get an idea of how to use CSV
; Consult the library (csv.ahk) for all available functions and required parameters
; and the results of each function

testCSV()

testCSV() {
    if FileExist("ExampleCSVFile.csv")
        FileDelete "ExampleCSVFile.csv"

    ; Creating an example CSV file
    FileAppend "
    (
        Year, Make, Model, Description, Price
        1997, Ford, E350, "ac, abs, moon", 3000.00
        1999, Chevy, "Venture " "Extended Edition" "", "", 4900.00
        1999, Chevy, "Venture " "Extended Edition, Very Large" "", , 5000.00
        1996, Jeep, Grand Cherokee, "MUST SELL! air, moon roof, loaded", 4799.00
    )", "ExampleCSVFile.csv"

    ; load a CSV file using CSV_Load(FileName, Delimiter)
    CSV_Load("ExampleCSVFile.csv")

    ; Reading a Cell using CSV_ReadCell(), should show "E350"
    MsgBox ("Contents of Cell in row number 2, and column 3: " CSV_ReadCell(2, 3))
}

' CalcToSQL v1.0.2
'
' MIT License
'
' Copyright (c) 2025 FarFromOkay
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Option Explicit
Option Compatible
Option Base 0 ' Explicitly set array base to 0

' --- Configuration Sheet Name ---
Const CONFIG_SHEET_NAME = "_CalcToSQL"

' Structure to hold column information
Type ColumnInfo
    CleanedHeader As String    ' Quoted, cleaned header name (`col`)
    OriginalIndex As Integer   ' Original column index in the sheet
    IsPK As Boolean            ' Derived from hints
    UniqueIndexNames As String ' Delimited list (e.g., ";u1;u2;")
    IndexNames As String       ' Delimited list (e.g., ";i1;i2;")
    InferredDataType As String ' SQL Data type inferred from data
    ' --- FK Fields ---
    IsFK As Boolean            ' Flag if this column is an FK
    FKRefTable As String       ' Referenced table name (quoted)
    FKRefColumn As String      ' Referenced column name (quoted)
    ' --- NEW: TSC Field ---
    IsTSC As Boolean           ' Flag if column is TIMESTAMP DEFAULT CURRENT_TIMESTAMP
End Type

' Structure to hold hint info from the config sheet
Type ConfigHintInfo
    SheetName As String
    ColumnName As String
    HintString As String
End Type


' ======= MAIN MACRO SUBROUTINE =======
Sub CalcToSQL()
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oSheet As Object
    Dim oCursor As Object
    Dim oCell As Object
    Dim oHeaderCell As Object
    Dim nLastRow As Long ' Last row potentially used (from cursor)
    Dim nTrueLastRow As Long ' Last row with actual data found
    Dim nLastCol As Long
    Dim nRow As Long
    Dim nCol As Long
    Dim i As Integer ' Sheet loop counter

    Dim sFilePath As String
    Dim nFileNum As Integer
    Dim sTableName As String          ' Quoted name of the current table
    Dim sUnquotedTableName As String  ' Unquoted name for generating constraint names
    Dim sColumnsPart As String ' For INSERT statement header list
    Dim sValuesPart As String  ' For INSERT statement values list
    Dim sSQL As String         ' Full INSERT/CREATE INDEX statement
    Dim sCreateTableSQL As String ' Full CREATE TABLE statement
    Dim sDropTableSQL As String   ' For per-sheet drop logic

    Dim aColumns() As ColumnInfo ' Array to hold info for valid columns

    Dim nValidColCount As Integer ' Columns per sheet
    Dim bNeedComma As Boolean
    Dim j As Long          ' Loop counter for columns (INSERT)
    Dim k As Long          ' Loop counter for columns (CREATE TABLE / Constraints)
    Dim pkIndex As Long    ' Loop counter for finding PKs
    Dim sPKClause As String ' String for PK list in CREATE TABLE
    Dim bFirstPK As Boolean ' Flag for comma in PK list / Constraint column list
    Dim nPKCount As Integer   ' Count of PK columns based on hints
    Dim nSinglePKColArrayIndex As Integer ' Array index if only one PK found

    Dim oNumFormats As Object ' For checking date formats

    ' Index/Hint Variables
    Dim allUniqueIndexNames As String
    Dim allIndexNames As String
    Dim currentHint As String
    Dim hintParts() As String
    Dim hintKey As String
    Dim hintValue As String
    Dim uniqueIndexName As String
    Dim indexName As String
    Dim constraintColumns As String
    Dim nameLoopIndex As Long
    Dim colLoopIndex As Long
    Dim distinctNames() As String
    Dim namePart As String

    ' FK parsing Variables
    Dim fkValue As String
    Dim fkOpenParen As Integer
    Dim fkCloseParen As Integer
    Dim fkRefTableRaw As String
    Dim fkRefColRaw As String

    Dim sheetName As String

    ' Counters for statistics
    Dim nSheetsProcessed As Long      : nSheetsProcessed = 0
    Dim nSheetsSkipped As Long        : nSheetsSkipped = 0
    Dim nTotalColumnsDefined As Long  : nTotalColumnsDefined = 0
    Dim nTotalRowsInserted As Long    : nTotalRowsInserted = 0 ' Count based on nTrueLastRow
    Dim nTotalUniqueKeys As Long      : nTotalUniqueKeys = 0
    Dim nTotalIndexes As Long         : nTotalIndexes = 0
    Dim nTotalForeignKeys As Long     : nTotalForeignKeys = 0
    Dim nTotalTSColumns As Long       : nTotalTSColumns = 0 ' Track TSC columns

    Dim sStatsMsg As String ' For final message box

    ' --- Configuration Settings ---
    Dim bDropIfExists As Boolean
    Dim vConfigHints As Variant ' To hold array from LoadConfigHints

    On Error GoTo ErrorHandler

    oDoc = ThisComponent
    If IsNull(oDoc) Or Not oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
        MsgBox "This macro must be run from a Calc spreadsheet document.", 16, "Error"
        Exit Sub
    End If

    oSheets = oDoc.getSheets()
    If oSheets.getCount() = 0 Then
        MsgBox "This document contains no sheets.", 48, "No Sheets Found"
        Exit Sub
    End If

    ' --- Load Configuration ---
    bDropIfExists = GetDropIfExistsSetting(oDoc) ' Default is FALSE
    vConfigHints = LoadConfigHints(oDoc)         ' Load hints from config sheet

    ' --- Get Output File Path ---
    sFilePath = GetSaveAsFilePath("Save SQL Create & Insert Statements As", "schema_inserts_" & Format(Now(), "YYYYMMDD_HHMMSS") & ".sql", "*.sql|SQL Files")
     If sFilePath = "" Then
        MsgBox "Operation cancelled by user.", 16, "Cancelled"
        Exit Sub
    End If

    ' --- Open File for Output ---
    nFileNum = FreeFile
    On Error Resume Next
    Open sFilePath For Output Lock Write As #nFileNum
    If Err <> 0 Then
         MsgBox "Error opening file for output!" & Chr(13) & "Path: '" & sFilePath & "'" & Chr(13) & "Error " & Err & ": " & Error$ & Chr(13) & Chr(13) & "Please ensure the directory exists and you have write permissions.", 16, "File Open Error"
         Close #nFileNum
        On Error GoTo ErrorHandler
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' --- Output Initial Setup ---
    Print #nFileNum, "-- SQL Script Generated by CalcToSQL v1.0.1" ' Updated version marker
    Print #nFileNum, "-- Source Document: " & oDoc.getTitle()
    Print #nFileNum, "-- Generated on: " & Format(Now(), "YYYY-MM-DD HH:MM:SS")
    Print #nFileNum, ""

    ' Disable Foreign Key Checks
    Print #nFileNum, "-- Disable foreign key checks."
    Print #nFileNum, "SET FOREIGN_KEY_CHECKS=0;"
    Print #nFileNum, ""

    ' --- Report DropIfExists Configuration ---
    If bDropIfExists Then
         Print #nFileNum, "-- Per-sheet 'DROP TABLE IF EXISTS' is enabled (DropIfExists=TRUE)."
    Else
         Print #nFileNum, "-- Per-sheet 'DROP TABLE IF EXISTS' is disabled (default or DropIfExists=FALSE/missing)."
    End If
    Print #nFileNum, ""


    oNumFormats = oDoc.getNumberFormats()

    ' --- Loop Through Each Sheet ---
    For i = 0 To oSheets.getCount() - 1
         oSheet = oSheets.getByIndex(i)
         sheetName = oSheet.getName()

         ' Check if sheet name starts with underscore (skip config sheet and others)
         If Left(sheetName, 1) = "_" Then
             If LCase(sheetName) <> LCase(CONFIG_SHEET_NAME) Then ' Don't log skipping the config sheet itself
                Print #nFileNum, "-- Skipping sheet '" & sheetName & "' (name starts with underscore)."
                Print #nFileNum, ""
             End If
             nSheetsSkipped = nSheetsSkipped + 1
             GoTo NextSheet
         End If

        nSheetsProcessed = nSheetsProcessed + 1

        Print #nFileNum, "-- Schema and statements for table derived from sheet: " & sheetName
        Print #nFileNum, "-- ======================================================="

        ' Determine used area (potential last row/col)
        oCursor = oSheet.createCursor()
        oCursor.gotoEndOfUsedArea(False)
        nLastRow = oCursor.getRangeAddress().EndRow
        nLastCol = oCursor.getRangeAddress().EndColumn

        ' Skip sheet if empty based on initial check
        If nLastRow < 0 Or nLastCol < 0 Then
             Print #nFileNum, "-- Sheet '" & sheetName & "' is empty or has no headers. Skipping."
             Print #nFileNum, ""
             GoTo NextSheet
        End If

        ' --- Get Table Name (Sheet Name) ---
        sUnquotedTableName = sheetName
        sTableName = "`" & Replace(sUnquotedTableName, "`", "``") & "`"

        ' --- Generate Per-Sheet Drop Statement (ONLY if enabled) ---
        If bDropIfExists Then
            sDropTableSQL = "DROP TABLE IF EXISTS " & sTableName & ";"
            Print #nFileNum, "-- Drop table if it exists (DropIfExists=TRUE)"
            Print #nFileNum, sDropTableSQL
            Print #nFileNum, "" ' Add a blank line
        End If

        ' --- Process Header Row (Row 0) ---
        sColumnsPart = ""
        nValidColCount = 0 ' Reset count for this sheet
        allUniqueIndexNames = ";"
        allIndexNames = ";"
        ReDim aColumns(0 To nLastCol) ' Initial dimensioning, will be preserved later
        bNeedComma = False ' For INSERT column list

        For nCol = 0 To nLastCol
            oHeaderCell = oSheet.getCellByPosition(nCol, 0)
            Dim rawHeaderText As String : rawHeaderText = Trim(oHeaderCell.getString())

            If rawHeaderText <> "" Then
                ' Check if raw header starts with underscore
                If Left(rawHeaderText, 1) = "_" Then
                    ' Skip column marked with leading underscore
                Else
                    Dim cleanedHeaderText As String : cleanedHeaderText = rawHeaderText
                    Dim unquotedCleanedHeaderText As String
                    Dim inlineHintString As String : inlineHintString = "" ' Hint from header [...]
                    Dim configHintString As String : configHintString = "" ' Hint from config sheet
                    Dim combinedHintString As String : combinedHintString = "" ' Merged hints
                    Dim openBracketPos As Integer : openBracketPos = InStr(rawHeaderText, "[")
                    Dim closeBracketPos As Integer : closeBracketPos = RevInStr(rawHeaderText, "]")

                    ' Extract inline hint
                    If openBracketPos > 0 And closeBracketPos > openBracketPos Then
                        cleanedHeaderText = Trim(Left(rawHeaderText, openBracketPos - 1))
                        inlineHintString = Mid(rawHeaderText, openBracketPos + 1, closeBracketPos - openBracketPos - 1)
                        inlineHintString = Trim(inlineHintString)
                    Else
                        cleanedHeaderText = rawHeaderText
                        inlineHintString = ""
                    End If

                    If cleanedHeaderText <> "" Then
                        unquotedCleanedHeaderText = cleanedHeaderText ' Store unquoted version for lookup

                        ' Find hint from config sheet (*** USING CORRECTED FUNCTION ***)
                        If IsArray(vConfigHints) Then ' Check if hints were loaded
                            configHintString = FindConfigHint(vConfigHints, sheetName, unquotedCleanedHeaderText)
                        End If

                        ' Combine and deduplicate hints
                        combinedHintString = CombineAndDeduplicateHints(inlineHintString, configHintString)

                        ' --- Set up column info ---
                        aColumns(nValidColCount).CleanedHeader = "`" & Replace(cleanedHeaderText, "`", "``") & "`"
                        aColumns(nValidColCount).OriginalIndex = nCol
                        aColumns(nValidColCount).IsPK = False
                        aColumns(nValidColCount).UniqueIndexNames = ";"
                        aColumns(nValidColCount).IndexNames = ";"
                        aColumns(nValidColCount).IsFK = False
                        aColumns(nValidColCount).FKRefTable = ""
                        aColumns(nValidColCount).FKRefColumn = ""
                        aColumns(nValidColCount).IsTSC = False ' Initialize TSC flag

                        ' Parse the COMBINED hint string
                        If combinedHintString <> "" Then
                            Dim hintsArr() As String : hintsArr = Split(combinedHintString, ",") ' Split combined hints
                            Dim hintIdx As Integer
                            For hintIdx = LBound(hintsArr) To UBound(hintsArr)
                                currentHint = Trim(LCase(hintsArr(hintIdx))) ' Use LCase for keys

                                Select Case currentHint
                                    Case "pk"
                                        aColumns(nValidColCount).IsPK = True
                                    Case "tsc" ' Check for TSC hint
                                        aColumns(nValidColCount).IsTSC = True
                                        nTotalTSColumns = nTotalTSColumns + 1 ' Increment counter
                                    Case Else
                                        ' Check for hints with values (uidx:, idx:, fk:)
                                        If InStr(currentHint, ":") > 0 Then
                                            hintParts = Split(currentHint, ":", 2)
                                            If UBound(hintParts) = 1 Then
                                                hintKey = Trim(hintParts(0))
                                                hintValue = Trim(hintParts(1)) ' Value can remain case-sensitive (e.g., table/col names)
                                                If hintValue <> "" Then
                                                    Select Case hintKey
                                                        Case "uidx"
                                                            uniqueIndexName = hintValue
                                                            If InStr(aColumns(nValidColCount).UniqueIndexNames, ";" & uniqueIndexName & ";") = 0 Then
                                                                aColumns(nValidColCount).UniqueIndexNames = aColumns(nValidColCount).UniqueIndexNames & uniqueIndexName & ";"
                                                            End If
                                                            If InStr(allUniqueIndexNames, ";" & uniqueIndexName & ";") = 0 Then
                                                                allUniqueIndexNames = allUniqueIndexNames & uniqueIndexName & ";"
                                                            End If
                                                        Case "idx"
                                                            indexName = hintValue
                                                            If InStr(aColumns(nValidColCount).IndexNames, ";" & indexName & ";") = 0 Then
                                                                aColumns(nValidColCount).IndexNames = aColumns(nValidColCount).IndexNames & indexName & ";"
                                                            End If
                                                            If InStr(allIndexNames, ";" & indexName & ";") = 0 Then
                                                                allIndexNames = allIndexNames & indexName & ";"
                                                            End If
                                                        Case "fk"
                                                            fkValue = hintValue
                                                            fkOpenParen = InStr(fkValue, "(")
                                                            fkCloseParen = InStr(fkValue, ")")
                                                            If fkOpenParen > 0 And fkCloseParen > fkOpenParen Then
                                                                fkRefTableRaw = Trim(Left(fkValue, fkOpenParen - 1))
                                                                fkRefColRaw = Trim(Mid(fkValue, fkOpenParen + 1, fkCloseParen - fkOpenParen - 1))
                                                                If fkRefTableRaw <> "" And fkRefColRaw <> "" Then
                                                                    aColumns(nValidColCount).IsFK = True
                                                                    aColumns(nValidColCount).FKRefTable = "`" & Replace(fkRefTableRaw, "`", "``") & "`"
                                                                    aColumns(nValidColCount).FKRefColumn = "`" & Replace(fkRefColRaw, "`", "``") & "`"
                                                                End If
                                                            End If
                                                    End Select
                                                End If
                                            End If
                                        End If ' End check for hints with values
                                End Select ' End Select Case currentHint
                            Next hintIdx
                        End If ' End parsing combinedHintString

                        ' Only add non-TSC columns to INSERT list
                        If Not aColumns(nValidColCount).IsTSC Then
                            If bNeedComma Then sColumnsPart = sColumnsPart & ", "
                            sColumnsPart = sColumnsPart & aColumns(nValidColCount).CleanedHeader
                            bNeedComma = True
                        End If

                        nValidColCount = nValidColCount + 1
                    End If ' End cleanedHeaderText <> ""
                End If ' End Left(rawHeaderText, 1) = "_"
            End If ' End rawHeaderText <> ""
        Next nCol ' Next Header Column

        If nValidColCount = 0 Then
            Print #nFileNum, "-- Sheet '" & sheetName & "' has no valid headers in Row 1. Skipping schema and data."
            Print #nFileNum, ""
            GoTo NextSheet
        End If

        nTotalColumnsDefined = nTotalColumnsDefined + nValidColCount
        ReDim Preserve aColumns(0 To nValidColCount - 1) ' Resize array to actual count

        ' Find the true last row with data by scanning backwards
        nTrueLastRow = 0 ' Default if no data rows found after header
        If nLastRow >= 1 Then ' Only search if potential rows exist
            For nRow = nLastRow To 1 Step -1
                Dim bFoundDataInRow As Boolean : bFoundDataInRow = False
                For j = 0 To nValidColCount - 1 ' Check relevant columns in this row
                    nCol = aColumns(j).OriginalIndex
                    oCell = oSheet.getCellByPosition(nCol, nRow)
                    If Trim(oCell.getString()) <> "" Then
                        bFoundDataInRow = True
                        Exit For
                    End If
                Next j
                If bFoundDataInRow Then
                    nTrueLastRow = nRow
                    Exit For
                End If
            Next nRow
        End If

        ' --- Infer Data Types (unless TSC) and Count PKs ---
        nPKCount = 0 : nSinglePKColArrayIndex = -1
        For k = 0 To nValidColCount - 1
            ' Only infer type if it's not a TSC column
            If Not aColumns(k).IsTSC Then
                ' *** CALL CORRECTED InferColumnType ***
                aColumns(k).InferredDataType = InferColumnType(oSheet, aColumns(k).OriginalIndex, 1, nTrueLastRow, oNumFormats)
            Else
                aColumns(k).InferredDataType = "TIMESTAMP DEFAULT CURRENT_TIMESTAMP" ' Set type directly for TSC
            End If

            If aColumns(k).IsPK Then
                nPKCount = nPKCount + 1
                If nPKCount = 1 Then nSinglePKColArrayIndex = k Else nSinglePKColArrayIndex = -1
            End If
        Next k

        ' --- Generate CREATE TABLE Statement ---
        sCreateTableSQL = "CREATE TABLE " & sTableName & " (" & Chr(10)
        bNeedComma = False ' Reset flag for separating items

        ' Part 1: Define Columns
        For k = 0 To nValidColCount - 1
            Dim currentDataType As String

            ' Handle TSC columns directly
            If aColumns(k).IsTSC Then
                currentDataType = "TIMESTAMP DEFAULT CURRENT_TIMESTAMP"
            Else
                ' Use inferred type, potentially adding AUTO_INCREMENT for single integer PK
                currentDataType = aColumns(k).InferredDataType
                If nPKCount = 1 And k = nSinglePKColArrayIndex Then
                     If UCase(currentDataType) = "INT UNSIGNED" Then
                        currentDataType = "INT UNSIGNED AUTO_INCREMENT"
                     ElseIf UCase(currentDataType) = "INT" Then
                        currentDataType = "INT AUTO_INCREMENT"
                     End If
                End If
            End If

            If bNeedComma Then sCreateTableSQL = sCreateTableSQL & "," & Chr(10)
            sCreateTableSQL = sCreateTableSQL & "  " & aColumns(k).CleanedHeader & " " & currentDataType
            bNeedComma = True
        Next k

        ' Part 2: Add PRIMARY KEY clause
        sPKClause = "" : bFirstPK = True
        For pkIndex = 0 To nValidColCount - 1
            If aColumns(pkIndex).IsPK Then
                If Not bFirstPK Then sPKClause = sPKClause & ", "
                sPKClause = sPKClause & aColumns(pkIndex).CleanedHeader
                bFirstPK = False
            End If
        Next pkIndex
        If sPKClause <> "" Then
             If bNeedComma Then sCreateTableSQL = sCreateTableSQL & "," & Chr(10)
             sCreateTableSQL = sCreateTableSQL & "  PRIMARY KEY (" & sPKClause & ")"
             bNeedComma = True
        End If

        ' Part 3: Add UNIQUE KEY constraints
        If allUniqueIndexNames <> ";" Then
            distinctNames = Split(Mid(allUniqueIndexNames, 2, Len(allUniqueIndexNames) - 2), ";")
            For nameLoopIndex = LBound(distinctNames) To UBound(distinctNames)
                uniqueIndexName = distinctNames(nameLoopIndex)
                If uniqueIndexName <> "" Then
                    constraintColumns = "" : bFirstPK = True
                    For colLoopIndex = 0 To nValidColCount - 1
                        If InStr(aColumns(colLoopIndex).UniqueIndexNames, ";" & uniqueIndexName & ";") > 0 Then
                            If Not bFirstPK Then constraintColumns = constraintColumns & ", "
                            constraintColumns = constraintColumns & aColumns(colLoopIndex).CleanedHeader
                            bFirstPK = False
                        End If
                    Next colLoopIndex
                    If constraintColumns <> "" Then
                        If bNeedComma Then sCreateTableSQL = sCreateTableSQL & "," & Chr(10)
                        Dim safeUQName As String : safeUQName = Replace(uniqueIndexName, "`", "``")
                        sCreateTableSQL = sCreateTableSQL & "  CONSTRAINT `uq_" & safeUQName & "` UNIQUE (" & constraintColumns & ")"
                        bNeedComma = True
                        nTotalUniqueKeys = nTotalUniqueKeys + 1
                    End If
                End If
            Next nameLoopIndex
        End If

        ' Part 4: Add FOREIGN KEY constraints
        For k = 0 To nValidColCount - 1
            If aColumns(k).IsFK Then
                If bNeedComma Then sCreateTableSQL = sCreateTableSQL & "," & Chr(10)
                Dim fkConstraintName As String
                Dim unquotedColName As String
                unquotedColName = Replace(aColumns(k).CleanedHeader, "`", "")
                fkConstraintName = "fk_" & sUnquotedTableName & "_" & unquotedColName
                fkConstraintName = Replace(fkConstraintName, "`", "``")
                ' Ensure FK name is reasonably short if needed (some DBs have limits)
                If Len(fkConstraintName) > 64 Then fkConstraintName = Left(fkConstraintName, 64) ' Example limit
                sCreateTableSQL = sCreateTableSQL & "  CONSTRAINT `" & fkConstraintName & "`" & _
                                  " FOREIGN KEY (" & aColumns(k).CleanedHeader & ")" & _
                                  " REFERENCES " & aColumns(k).FKRefTable & " (" & aColumns(k).FKRefColumn & ")"
                bNeedComma = True
                nTotalForeignKeys = nTotalForeignKeys + 1
            End If
        Next k

        ' Part 5: Close CREATE TABLE statement
        sCreateTableSQL = sCreateTableSQL & Chr(10) & ");"

        ' Print CREATE TABLE statement
        Print #nFileNum, "-- Create table structure for " & sTableName
        Print #nFileNum, sCreateTableSQL
        Print #nFileNum, ""

        ' Part 6: Add CREATE INDEX statements (after table)
        If allIndexNames <> ";" Then
            Print #nFileNum, "-- Create non-unique indexes for " & sTableName
            distinctNames = Split(Mid(allIndexNames, 2, Len(allIndexNames) - 2), ";")
            For nameLoopIndex = LBound(distinctNames) To UBound(distinctNames)
                indexName = distinctNames(nameLoopIndex)
                If indexName <> "" Then
                    constraintColumns = "" : bFirstPK = True
                    For colLoopIndex = 0 To nValidColCount - 1
                        If InStr(aColumns(colLoopIndex).IndexNames, ";" & indexName & ";") > 0 Then
                            If Not bFirstPK Then constraintColumns = constraintColumns & ", "
                            constraintColumns = constraintColumns & aColumns(colLoopIndex).CleanedHeader
                            bFirstPK = False
                        End If
                    Next colLoopIndex
                    If constraintColumns <> "" Then
                        Dim safeIdxName As String : safeIdxName = Replace(indexName, "`", "``")
                        sSQL = "CREATE INDEX `idx_" & safeIdxName & "` ON " & sTableName & " (" & constraintColumns & ");"
                        Print #nFileNum, sSQL
                        nTotalIndexes = nTotalIndexes + 1
                    End If
                End If
            Next nameLoopIndex
            Print #nFileNum, ""
        End If


        ' Check if there are data rows before processing them
        If nTrueLastRow < 1 Then
              Print #nFileNum, "-- Sheet '" & sheetName & "' has headers but no data rows with content. Skipping INSERT statements."
              Print #nFileNum, ""
             GoTo NextSheet
        End If

        ' Check if there are any columns to insert (after excluding TSC)
        If Trim(sColumnsPart) = "" Then
             Print #nFileNum, "-- No columns eligible for INSERT (all might be TSC or skipped). Skipping INSERT statements."
             Print #nFileNum, ""
             GoTo NextSheet ' Skip to next sheet if no columns to insert
        End If

        ' Process Data Rows for INSERT Statements
        Print #nFileNum, "-- Data for table " & sTableName
        For nRow = 1 To nTrueLastRow
            sValuesPart = ""
            bNeedComma = False ' Reset for value list

            ' Only include values for non-TSC columns
            For j = 0 To nValidColCount - 1
                ' Check if this column should be included in INSERTs
                If Not aColumns(j).IsTSC Then
                    nCol = aColumns(j).OriginalIndex
                    oCell = oSheet.getCellByPosition(nCol, nRow)
                     If bNeedComma Then sValuesPart = sValuesPart & ", "
                    ' *** CALL CORRECTED FormatSQLValue (v4) ***
                    sValuesPart = sValuesPart & FormatSQLValue(oCell, oNumFormats)
                    bNeedComma = True
                End If
            Next j

            ' Only print INSERT if there are values (should match sColumnsPart check)
            If Trim(sValuesPart) <> "" Then
                sSQL = "INSERT INTO " & sTableName & " (" & sColumnsPart & ") VALUES (" & sValuesPart & ");"
                Print #nFileNum, sSQL
                nTotalRowsInserted = nTotalRowsInserted + 1
            End If
        Next nRow

        Print #nFileNum, "" ' Add a blank line after processing a sheet's data

NextSheet: ' Label for skipping sheet processing
        Print #nFileNum, "" ' Extra blank line between sheets
    Next i ' Next Sheet

    ' --- Clean Up ---
    ' Re-enable Foreign Key Checks at the end of the script
    Print #nFileNum, "" ' Add a blank line before the final command
    Print #nFileNum, "-- Re-enable foreign key checks."
    Print #nFileNum, "SET FOREIGN_KEY_CHECKS=1;"

    Close #nFileNum

    ' Build statistics message
    sStatsMsg = "SQL statements generated successfully." & Chr(13) & _
                "Saved to: " & sFilePath & Chr(13) & Chr(13) & _
                "Statistics:" & Chr(13) & _
                "------------" & Chr(13) & _
                "Sheets Processed (Tables): " & nSheetsProcessed & Chr(13) & _
                "Sheets Skipped (_): " & nSheetsSkipped & Chr(13) & _
                "Total Columns Defined: " & nTotalColumnsDefined & Chr(13) & _
                "Timestamp Columns (TSC): " & nTotalTSColumns & Chr(13) & _
                "Total Rows Inserted: " & nTotalRowsInserted & Chr(13) & _
                "Unique Keys Created: " & nTotalUniqueKeys & Chr(13) & _
                "Foreign Keys Created: " & nTotalForeignKeys & Chr(13) & _
                "Standard Indexes Created: " & nTotalIndexes

    MsgBox sStatsMsg, 64, "Complete" ' Use vbInformation icon (64)

    Exit Sub ' Normal exit

ErrorHandler:
    MsgBox "An error occurred:" & Chr(13) & "Error " & Err & ": " & Error$ & Chr(13) & "On or near line " & Erl, 16, "Macro Error"
    If nFileNum > 0 Then
        On Error Resume Next
        ' Attempt to re-enable checks even on error, before closing
        Print #nFileNum, ""
        Print #nFileNum, "-- Attempting to re-enable foreign key checks after error."
        Print #nFileNum, "SET FOREIGN_KEY_CHECKS=1;"
        Close #nFileNum
        On Error GoTo 0
    End If
End Sub


' --- Helper Function to Get DropIfExists Setting ---
' Checks the config sheet for a 'DropIfExists' setting within a 'Settings' section.
' Returns TRUE only if explicitly set to TRUE/Yes/1/On.
' Returns FALSE by default, or if setting/section is missing, or if set to FALSE/No/0/Off.
Function GetDropIfExistsSetting(oDoc As Object) As Boolean
    Dim oConfigSheet As Object
    Dim oCell As Object
    Dim oCellKey As Object      ' Cell object for the key (Column A)
    Dim oCellValue As Object    ' Cell object for the value (Column B)
    Dim sCellValue As String    ' Value read from Column A (used for section/key finding)
    Dim sKey As String          ' Key read from Column A within section
    Dim sValue As String        ' Value read from Column B within section
    Dim lCaseValue As String    ' Lowercase version of sValue
    Dim nRow As Long            ' Row index (0-based)
    Dim nMaxRow As Long         ' To avoid checking entire sheet
    Dim errNum As Long          ' To store error number
    Dim foundSectionHeader As Boolean
    Dim settingsDataStartRow As Long

    ' *** Default to FALSE (Do NOT drop unless explicitly told to) ***
    GetDropIfExistsSetting = FALSE
    foundSectionHeader = False
    settingsDataStartRow = -1

    On Error Resume Next ' Ignore errors if sheet doesn't exist
    oConfigSheet = oDoc.getSheets().getByName(CONFIG_SHEET_NAME)
    errNum = Err.Number ' Store Err.Number immediately
    On Error GoTo 0 ' Reset error handling NOW

    If errNum <> 0 Then
        ' Config sheet not found or error accessing it. Return default FALSE.
        Exit Function
    End If
    ' If we get here, errNum was 0, sheet should be valid

    ' Find the last used row in the config sheet to limit scanning
    Dim oCursor As Object
    oCursor = oConfigSheet.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    nMaxRow = oCursor.getRangeAddress().EndRow
    If nMaxRow < 0 Then nMaxRow = 0 ' Handle empty sheet case
    If nMaxRow > 500 Then nMaxRow = 500 ' Safety limit

    ' --- First Pass: Find the start of the Settings data ---
    For nRow = 0 To nMaxRow
        oCell = oConfigSheet.getCellByPosition(0, nRow) ' Column A
        sCellValue = Trim(oCell.getString())
        If LCase(sCellValue) = "settings" Then
             foundSectionHeader = True
             ' Assuming Header("Settings") at index nRow, Blank at nRow+1, SubHeader("Name") at nRow+2, Data starts at nRow+3
             If nRow + 3 <= nMaxRow Then ' Check bounds
                 settingsDataStartRow = nRow + 3
             Else
                 settingsDataStartRow = -1 ' Not enough rows below header for data
             End If
             Exit For ' Found the section header row
        End If
    Next nRow

    ' --- Second Pass: If section found, parse it ---
    If Not foundSectionHeader Or settingsDataStartRow < 0 Then
        Exit Function ' Section not found, or no room for data, return default FALSE
    End If

    nRow = settingsDataStartRow ' Start from the calculated data row
    Do
        oCellKey = oConfigSheet.getCellByPosition(0, nRow) ' Column A (Key)
        sKey = Trim(oCellKey.getString())

        If sKey = "" Then
            ' Blank row marks end of section data
            Exit Do
        End If

        ' Check if this is the key we are looking for (case-insensitive)
        If LCase(sKey) = "dropifexists" Then
            oCellValue = oConfigSheet.getCellByPosition(1, nRow) ' Column B (Value)
            sValue = Trim(oCellValue.getString())
            lCaseValue = LCase(sValue)

            ' Check if the value represents TRUE
            If lCaseValue = "true" Or sValue = "1" Or lCaseValue = "yes" Or lCaseValue = "on" Then
                GetDropIfExistsSetting = TRUE
            Else
                GetDropIfExistsSetting = FALSE ' Explicitly false if key found but value isn't true
            End If
            Exit Function ' Found our key, exit function immediately with the determined value
        End If

        nRow = nRow + 1
        If nRow > nMaxRow Then Exit Do ' Safety break if max row reached
    Loop

    ' Return default FALSE if loop finished without finding the key within the section
End Function

' Function to load hints from the config sheet into an array
Function LoadConfigHints(oDoc As Object) As Variant
    Dim oConfigSheet As Object
    Dim oCell As Object
    Dim sCellValue As String
    Dim nRow As Long
    Dim nMaxRow As Long
    Dim errNum As Long
    Dim foundSectionHeader As Boolean
    Dim hintsDataStartRow As Long
    Dim hintsCount As Long
    Dim aHints() As ConfigHintInfo ' Dynamic array to hold hints
    Dim currentHint As ConfigHintInfo

    ' Initialize return value
    LoadConfigHints = Empty ' Return Empty if no hints found or error
    foundSectionHeader = False
    hintsDataStartRow = -1
    hintsCount = 0

    On Error Resume Next ' Ignore errors if sheet doesn't exist
    oConfigSheet = oDoc.getSheets().getByName(CONFIG_SHEET_NAME)
    errNum = Err.Number
    On Error GoTo 0 ' Reset error handling

    If errNum <> 0 Then Exit Function ' Config sheet not found

    ' Find the last used row
    Dim oCursor As Object
    oCursor = oConfigSheet.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    nMaxRow = oCursor.getRangeAddress().EndRow
    If nMaxRow < 0 Then nMaxRow = 0
    If nMaxRow > 1000 Then nMaxRow = 1000 ' Safety limit for hint section

    ' --- First Pass: Find the start of the Hints section ---
    For nRow = 0 To nMaxRow
        oCell = oConfigSheet.getCellByPosition(0, nRow) ' Column A
        sCellValue = Trim(oCell.getString())
        If LCase(sCellValue) = "hints" Then
             foundSectionHeader = True
             ' Assuming Header("Hints") at nRow, Blank at nRow+1, SubHeaders("Sheet","Column","Hint") at nRow+2, Data starts at nRow+3
             If nRow + 3 <= nMaxRow Then ' Check bounds
                 ' Basic validation of expected headers
                 Dim headerSheet As String : headerSheet = Trim(oConfigSheet.getCellByPosition(0, nRow + 2).getString())
                 Dim headerCol As String   : headerCol = Trim(oConfigSheet.getCellByPosition(1, nRow + 2).getString())
                 Dim headerHint As String  : headerHint = Trim(oConfigSheet.getCellByPosition(2, nRow + 2).getString())
                 If LCase(headerSheet) = "sheet" And LCase(headerCol) = "column" And LCase(headerHint) = "hint" Then
                     hintsDataStartRow = nRow + 3
                 Else
                     hintsDataStartRow = -1 ' Headers don't match expected
                 End If
             Else
                 hintsDataStartRow = -1 ' Not enough rows below header for data
             End If
             Exit For ' Found the section header row
        End If
    Next nRow

    ' --- Second Pass: If section and headers found, parse hints ---
    If Not foundSectionHeader Or hintsDataStartRow < 0 Then
        Exit Function ' Section or headers not found, return Empty
    End If

    nRow = hintsDataStartRow ' Start from the calculated data row
    Do
        Dim sheetVal As String : sheetVal = Trim(oConfigSheet.getCellByPosition(0, nRow).getString())
        Dim colVal As String   : colVal = Trim(oConfigSheet.getCellByPosition(1, nRow).getString())
        Dim hintVal As String  : hintVal = Trim(oConfigSheet.getCellByPosition(2, nRow).getString())

        If sheetVal = "" Then
            ' Blank row in Sheet column marks end of section data
            Exit Do
        End If

        ' Only add if Sheet, Column, and Hint are provided
        If colVal <> "" And hintVal <> "" Then
            hintsCount = hintsCount + 1
            ReDim Preserve aHints(0 To hintsCount - 1) ' Resize array

            ' Store data in the new element
            aHints(hintsCount - 1).SheetName = sheetVal
            aHints(hintsCount - 1).ColumnName = colVal
            aHints(hintsCount - 1).HintString = hintVal
        End If

        nRow = nRow + 1
        If nRow > nMaxRow Then Exit Do ' Safety break if max row reached
    Loop

    ' Return the array if hints were found
    If hintsCount > 0 Then
        LoadConfigHints = aHints
    End If
    ' Otherwise, it remains Empty
End Function


' *** BEGIN CORRECTED FindConfigHint FUNCTION ***
' Helper Function to find ALL hints in the config array for a given sheet/column
' and combine them with commas.
Function FindConfigHint(aHints As Variant, sSheetName As String, sColumnName As String) As String
    Dim i As Long
    Dim lowerSheetName As String : lowerSheetName = LCase(sSheetName)
    Dim lowerColName As String   : lowerColName = LCase(sColumnName)
    Dim combinedHints As String  : combinedHints = "" ' Initialize empty string
    Dim needsComma As Boolean    : needsComma = False

    ' Default: empty string if no hints found
    FindConfigHint = ""

    If IsEmpty(aHints) Or Not IsArray(aHints) Then Exit Function ' Exit if input is not a valid array

    For i = LBound(aHints) To UBound(aHints)
        ' Case-insensitive comparison for sheet and column names
        If LCase(aHints(i).SheetName) = lowerSheetName And LCase(aHints(i).ColumnName) = lowerColName Then
            ' Found a matching hint row
            If Trim(aHints(i).HintString) <> "" Then ' Only add non-empty hints
                If needsComma Then
                    combinedHints = combinedHints & "," ' Add comma before subsequent hints
                End If
                combinedHints = combinedHints & Trim(aHints(i).HintString) ' Append the hint
                needsComma = True ' Set flag to add comma next time
            End If
            ' --- DO NOT Exit Function here - continue searching for more matches ---
        End If
    Next i

    ' Return the combined string of all found hints
    FindConfigHint = combinedHints

End Function
' *** END CORRECTED FindConfigHint FUNCTION ***


' Helper Function to combine two hint strings and remove duplicates
Function CombineAndDeduplicateHints(hint1 As String, hint2 As String) As String
    Dim combined As String
    Dim parts1() As String
    Dim parts2() As String
    Dim finalHints As String: finalHints = ";" ' Use semicolon delimiter for easy InStr check
    Dim part As Variant ' Loop variable for arrays
    Dim trimmedPart As String

    ' Combine initial strings with a comma if both exist
    If Trim(hint1) <> "" And Trim(hint2) <> "" Then
        combined = Trim(hint1) & "," & Trim(hint2)
    ElseIf Trim(hint1) <> "" Then
        combined = Trim(hint1)
    ElseIf Trim(hint2) <> "" Then
        combined = Trim(hint2)
    Else
        CombineAndDeduplicateHints = "" ' Both empty
        Exit Function
    End If

    ' Split the combined string
    parts1 = Split(combined, ",")

    ' Process each part, adding unique ones to finalHints
    For Each part In parts1
        trimmedPart = Trim(LCase(part)) ' Use lowercase for comparison
        If trimmedPart <> "" Then
            ' Check if this part (with delimiters) is already in the final string
            If InStr(finalHints, ";" & trimmedPart & ";") = 0 Then
                finalHints = finalHints & trimmedPart & ";"
            End If
        End If
    Next part

    ' Convert the semicolon-delimited string back to comma-delimited
    If Len(finalHints) > 1 Then ' More than just the initial ";"
        finalHints = Mid(finalHints, 2, Len(finalHints) - 2) ' Remove leading/trailing ";"
        CombineAndDeduplicateHints = Replace(finalHints, ";", ",")
    Else
        CombineAndDeduplicateHints = "" ' No valid hints found
    End If

End Function


' --- Helper Function to Infer SQL Data Type by Scanning Column Data ---
' NOTE: This function is NOT called for columns marked with the TSC hint.
' CORRECTED VERSION: Prioritizes getString() for formulas to avoid misinterpreting string results as numbers.
Function InferColumnType(oSheet As Object, nColIndex As Integer, nStartDataRow As Integer, nTrueEndRow As Long, oNumFormats As Object) As String
    Dim nRow As Long
    Dim oCell As Object
    Dim cellType As Integer
    Dim val As Variant
    Dim sCellValue As String ' Added to store string value directly
    Dim nFormatId As Long
    Dim oFormatProps As Object

    Dim hasText As Boolean     : hasText = False
    Dim hasDateTime As Boolean : hasDateTime = False
    Dim hasDecimal As Boolean  : hasDecimal = False
    Dim hasInteger As Boolean  : hasInteger = False
    Dim hasNegativeInteger As Boolean : hasNegativeInteger = False ' Flag for INT vs INT UNSIGNED
    Dim onlyEmpty As Boolean   : onlyEmpty = True

    If nTrueEndRow < nStartDataRow Then nTrueEndRow = nStartDataRow ' Ensure at least one row is checked if header exists

    For nRow = nStartDataRow To nTrueEndRow
        oCell = oSheet.getCellByPosition(nColIndex, nRow)
        cellType = oCell.getType()

        If cellType <> com.sun.star.table.CellContentType.EMPTY Then
            onlyEmpty = False

            Select Case cellType
                Case com.sun.star.table.CellContentType.TEXT
                    hasText = True
                    Exit For ' Text overrides everything else

                Case com.sun.star.table.CellContentType.VALUE
                    val = oCell.getValue()
                    nFormatId = oCell.NumberFormat
                    On Error Resume Next
                    oFormatProps = oNumFormats.getByKey(nFormatId)
                    On Error GoTo 0 ' Reset error handling

                    If Not IsEmpty(oFormatProps) And Not IsNull(oFormatProps) Then
                        If (oFormatProps.Type AND com.sun.star.util.NumberFormat.DATE) <> 0 OR _
                           (oFormatProps.Type AND com.sun.star.util.NumberFormat.DATETIME) <> 0 OR _
                           (oFormatProps.Type AND com.sun.star.util.NumberFormat.TIME) <> 0 Then
                            hasDateTime = True
                        Else
                            If IsNumeric(val) Then ' Check if it's actually numeric before Fix
                                If val <> Fix(val) Then
                                    hasDecimal = True
                                Else
                                    hasInteger = True
                                    If val < 0 Then hasNegativeInteger = True
                                End If
                            Else ' Treat non-numeric value as text if format wasn't date/time
                                hasText = True : Exit For
                            End If
                        End If
                    Else ' No format info available
                         If IsNumeric(val) Then ' Check if it's numeric before Fix
                             If val <> Fix(val) Then
                                 hasDecimal = True
                             Else
                                 hasInteger = True
                                 If val < 0 Then hasNegativeInteger = True
                             End If
                         Else ' Treat non-numeric value as text if no format info
                             hasText = True : Exit For
                         End If
                    End If

                Case com.sun.star.table.CellContentType.FORMULA
                    If oCell.getError() = 0 Then
                        val = oCell.getValue() ' Get numeric/date value if applicable
                        sCellValue = Trim(oCell.getString()) ' Get string representation

                        If sCellValue = "" Then
                            ' Skip empty string results, don't change type inference
                        Else
                            ' --- Prioritize String Check ---
                            ' Check if the string result *looks* like a number or date.
                            ' If not, it's definitely text.
                            If Not IsNumeric(sCellValue) And Not IsDate(sCellValue) Then
                                 hasText = True
                                 Exit For ' Found text, no need to check further in this column
                            Else
                                ' String result looks numeric or date-like. Now check format/value for specifics.
                                nFormatId = oCell.NumberFormat
                                On Error Resume Next
                                oFormatProps = oNumFormats.getByKey(nFormatId)
                                On Error GoTo 0 ' Reset error handling

                                Dim isDateType As Boolean : isDateType = False
                                Dim isNumericType As Boolean : isNumericType = False

                                If Not IsEmpty(oFormatProps) And Not IsNull(oFormatProps) Then
                                    If (oFormatProps.Type AND com.sun.star.util.NumberFormat.DATE) <> 0 OR _
                                       (oFormatProps.Type AND com.sun.star.util.NumberFormat.DATETIME) <> 0 OR _
                                       (oFormatProps.Type AND com.sun.star.util.NumberFormat.TIME) <> 0 Then
                                        isDateType = True
                                    ElseIf (oFormatProps.Type AND com.sun.star.util.NumberFormat.NUMBER) <> 0 OR _
                                           (oFormatProps.Type AND com.sun.star.util.NumberFormat.CURRENCY) <> 0 OR _
                                           (oFormatProps.Type AND com.sun.star.util.NumberFormat.PERCENT) <> 0 OR _
                                           (oFormatProps.Type AND com.sun.star.util.NumberFormat.FRACTION) <> 0 Then
                                        isNumericType = True
                                    ' Else: Format is General, Text, etc. - rely on value check below
                                    End If
                                End If ' End format check

                                ' Refine based on format or value if format wasn't decisive
                                If isDateType Then
                                     hasDateTime = True
                                ElseIf isNumericType Then ' Format explicitly numeric
                                     If val <> Fix(val) Then hasDecimal = True Else hasInteger = True
                                     If val < 0 Then hasNegativeInteger = True
                                Else ' Format wasn't decisive (or no format info) - check value type
                                     If IsDate(val) Then ' Check underlying value first
                                         hasDateTime = True
                                     ElseIf IsNumeric(val) Then ' Check underlying value
                                         If val <> Fix(val) Then hasDecimal = True Else hasInteger = True
                                         If val < 0 Then hasNegativeInteger = True
                                     Else
                                         ' If string looked numeric/date but value isn't, treat as text (fallback)
                                         hasText = True : Exit For
                                     End If
                                End If
                            End If ' End check: If Not IsNumeric(sCellValue)...
                        End If ' End check: sCellValue = ""
                    Else ' Formula resulted in an error
                         hasText = True : Exit For
                    End If
            End Select ' End Select cellType
        End If ' End If cellType <> EMPTY
        If hasText Then Exit For ' Exit outer loop if text found
    Next nRow

    ' Determine final type based on flags (priority: Text > DateTime > Decimal > Integer)
    If onlyEmpty Then
        InferColumnType = "VARCHAR(255)" ' Default for empty columns
    ElseIf hasText Then
        InferColumnType = "VARCHAR(255)" ' Configurable length might be better
    ElseIf hasDateTime Then
        InferColumnType = "DATETIME"
    ElseIf hasDecimal Then
        InferColumnType = "DOUBLE" ' Or DECIMAL(p,s) if more precision needed
    ElseIf hasInteger Then
        If hasNegativeInteger Then
            InferColumnType = "INT"
        Else
            InferColumnType = "INT UNSIGNED"
        End If
    Else
        InferColumnType = "VARCHAR(255)" ' Fallback if only empty cells encountered but onlyEmpty wasn't true (shouldn't happen)
    End If
End Function


' --- Helper Function to Format Cell Value for SQL ---
' CORRECTED VERSION (v4): Stronger prioritization of getString() for formulas.
Function FormatSQLValue(oCell As Object, oNumFormats As Object) As String
    Dim sResult As String
    Dim nFormatId As Long
    Dim oFormatProps As Object
    Dim val As Variant
    Dim sCellValue As String

    On Error GoTo FormatErrorHandler

    Select Case oCell.getType()
        Case com.sun.star.table.CellContentType.EMPTY
            sResult = "NULL"

        Case com.sun.star.table.CellContentType.VALUE
            val = oCell.getValue()
            nFormatId = oCell.NumberFormat
            On Error Resume Next
            oFormatProps = oNumFormats.getByKey(nFormatId)
            On Error GoTo FormatErrorHandler ' Reset error handling

            If Not IsEmpty(oFormatProps) And Not IsNull(oFormatProps) Then
                  If (oFormatProps.Type AND com.sun.star.util.NumberFormat.DATE) <> 0 OR _
                    (oFormatProps.Type AND com.sun.star.util.NumberFormat.DATETIME) <> 0 OR _
                    (oFormatProps.Type AND com.sun.star.util.NumberFormat.TIME) <> 0 Then
                     ' Format date/time consistently for SQL
                     sResult = "'" & Format(val, "YYYY-MM-DD HH:MM:SS") & "'"
                 ElseIf IsNumeric(val) Then ' Explicitly check if numeric
                     ' Use Str() for numbers to avoid locale issues with decimal points
                     sResult = Trim(Str(val))
                     ' Basic standardisation - ensure '.' is decimal separator if needed
                     sResult = Replace(sResult, ",", ".")
                 Else ' Treat as text if format is not date/time and value isn't numeric
                     sCellValue = oCell.getString()
                     sResult = "'" & Replace(sCellValue, "'", "''") & "'" ' Escape single quotes
                 End If
            Else ' Fallback if format info unavailable
                 If IsDate(val) Then ' Check if it looks like a date
                     sResult = "'" & Format(val, "YYYY-MM-DD HH:MM:SS") & "'"
                 ElseIf IsNumeric(val) Then ' Check if numeric
                     sResult = Trim(Str(val))
                     sResult = Replace(sResult, ",", ".")
                 Else ' Otherwise treat as text
                     sCellValue = oCell.getString()
                     sResult = "'" & Replace(sCellValue, "'", "''") & "'"
                 End If
            End If

        Case com.sun.star.table.CellContentType.TEXT
            sCellValue = oCell.getString()
            If UCase(Trim(sCellValue)) = "NULL" Then ' Allow explicit NULL keyword in text cells
                sResult = "NULL"
            Else
                sResult = "'" & Replace(sCellValue, "'", "''") & "'" ' Escape single quotes
            End If

        Case com.sun.star.table.CellContentType.FORMULA
            ' --- Start v4 Correction for FORMULA ---
            If oCell.getError() <> 0 Then
                 sResult = "NULL" ' Treat formula errors as NULL
            Else
                ' Always get the string value first for formulas
                sCellValue = oCell.getString()

                ' Handle empty or explicit "NULL" string results
                If Trim(sCellValue) = "" Then
                    sResult = "NULL"
                ElseIf UCase(Trim(sCellValue)) = "NULL" Then
                    sResult = "NULL"
                Else
                    ' Get underlying value and format info
                    val = oCell.getValue()
                    nFormatId = oCell.NumberFormat
                    Dim formatIsDateTime As Boolean : formatIsDateTime = False
                    Dim formatIsNumeric As Boolean : formatIsNumeric = False

                    On Error Resume Next
                    oFormatProps = oNumFormats.getByKey(nFormatId)
                    If Err.Number = 0 And Not IsEmpty(oFormatProps) And Not IsNull(oFormatProps) Then
                         ' Check format type for Date/Time
                         If (oFormatProps.Type AND com.sun.star.util.NumberFormat.DATE) <> 0 OR _
                           (oFormatProps.Type AND com.sun.star.util.NumberFormat.DATETIME) <> 0 OR _
                           (oFormatProps.Type AND com.sun.star.util.NumberFormat.TIME) <> 0 Then
                            formatIsDateTime = True
                         ' Check if format is explicitly numeric (excluding boolean, scientific etc.)
                         ElseIf (oFormatProps.Type AND com.sun.star.util.NumberFormat.NUMBER) <> 0 OR _
                                (oFormatProps.Type AND com.sun.star.util.NumberFormat.CURRENCY) <> 0 OR _
                                (oFormatProps.Type AND com.sun.star.util.NumberFormat.PERCENT) <> 0 OR _
                                (oFormatProps.Type AND com.sun.star.util.NumberFormat.FRACTION) <> 0 Then
                            formatIsNumeric = True
                         End If
                    End If
                    On Error GoTo FormatErrorHandler ' Re-enable default handler

                    ' Decision Logic:
                    ' 1. If format is Date/Time AND value is date -> Use formatted date value
                    ' 2. Else If format is Numeric AND value is numeric AND string matches numeric -> Use numeric value
                    ' 3. Else (Default) -> Use string value
                    If formatIsDateTime And IsDate(val) Then
                        sResult = "'" & Format(val, "YYYY-MM-DD HH:MM:SS") & "'"
                    ElseIf formatIsNumeric And IsNumeric(val) And (Trim(Str(val)) = sCellValue) Then
                        ' Only format as number if the string representation exactly matches the numeric value
                        sResult = Trim(Str(val))
                        sResult = Replace(sResult, ",", ".")
                    Else
                        ' Default to using the string value for General format, Text format,
                        ' or when numeric value/string representation mismatch (e.g., "player1" vs 0)
                        sResult = "'" & Replace(sCellValue, "'", "''") & "'"
                    End If
                End If ' End check for Trim(sCellValue) = "" or "NULL"
            End If ' End check oCell.getError() <> 0
            ' --- End v4 Correction for FORMULA ---

        Case Else
            ' Any other cell type (e.g., Error) treat as NULL
            sResult = "NULL"
    End Select

    FormatSQLValue = sResult
    Exit Function
FormatErrorHandler:
    ' Return a distinct error string if formatting fails
    FormatSQLValue = "'SQL_FMT_ERR'"
End Function


' --- Helper Function to Show Save As Dialog ---
Function GetSaveAsFilePath(sTitle As String, Optional sInitialFilename As String, Optional sFilter As String) As String
    Dim oFilePicker As Object
    Dim sResultPath As String
    On Error Resume Next
    oFilePicker = CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    If Err <> 0 Then
        MsgBox "Error creating FilePicker service: " & Error$, 16, "Error"
        GetSaveAsFilePath = ""
        Exit Function
    End If
    On Error GoTo 0 ' Reset error handling

    oFilePicker.setTitle(sTitle)
    oFilePicker.setMultiSelectionMode(False)
    If sInitialFilename <> "" Then oFilePicker.setDefaultName(sInitialFilename)

    If sFilter <> "" Then
        Dim filterParts() As String
        filterParts = Split(sFilter, "|")
        If UBound(filterParts) = 1 Then
            oFilePicker.appendFilter(filterParts(1), filterParts(0))
            oFilePicker.setCurrentFilter(filterParts(1))
        Else
             oFilePicker.appendFilter("All Files (*.*)", "*.*")
        End If
    Else
         oFilePicker.appendFilter("All Files (*.*)", "*.*")
    End If

    oFilePicker.Initialize(Array(com.sun.star.ui.dialogs.TemplateDescription.FILESAVE_SIMPLE))

    If oFilePicker.execute() = com.sun.star.ui.dialogs.ExecutableDialogResults.OK Then
        Dim aFiles As Variant : aFiles = oFilePicker.getFiles()
        If UBound(aFiles) >= 0 Then
            sResultPath = ConvertFromURL(aFiles(0)) ' Convert file URL to system path
        Else
            sResultPath = ""
        End If
    Else
        sResultPath = ""
    End If

    GetSaveAsFilePath = sResultPath
    oFilePicker = Nothing
End Function

' --- Helper function to find last occurrence of a character ---
' Needed because Basic doesn't have a built-in InStrRev
Function RevInStr(sLookIn As String, sLookFor As String) As Integer
    Dim i As Integer
    Dim foundPos As Integer : foundPos = 0
    For i = 1 To Len(sLookIn)
        If Mid(sLookIn, i, Len(sLookFor)) = sLookFor Then foundPos = i
    Next i
    RevInStr = foundPos
End Function
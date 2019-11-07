Attribute VB_Name = "QuickfsParserModule"
Option Explicit
Option Private Module

' Locate all QFS formulas in a string and evaluate required keys for each
' If formula can be resolved to a static value, this is returned
Function ParseFormula(Formula As String, cell As Range, sheet As Worksheet, ByRef keys)
    Dim fn As String: fn = ""
    Dim resolved As String: resolved = ""
    Dim resolvable As Boolean: resolvable = True
    Dim quotes As Boolean: quotes = False
    Dim inQFS As Long: inQFS = 0
    Dim hasQFS As Boolean: hasQFS = False
    Dim parens As Long: parens = 0
    Dim i As Long
    For i = 1 To VBA.Len(Formula)
        Dim char As String
        char = VBA.Mid(Formula, i, 1)
        resolved = resolved & char
        If char = """" Then
            quotes = Not quotes
            If VBA.Len(fn) > 0 Then
                fn = fn & char
            End If
        ElseIf quotes Then
            If VBA.Len(fn) > 0 Then
                fn = fn & char
            End If
        ElseIf parens = 0 And inQFS = 0 And (char = "Q" Or char = "q") Then
            fn = fn & char
            inQFS = 1
        ElseIf parens = 0 And inQFS = 1 And (char = "F" Or char = "f") Then
            fn = fn & char
            inQFS = 2
        ElseIf parens = 0 And inQFS = 2 And (char = "S" Or char = "s") Then
            fn = fn & char
            inQFS = 3
        ElseIf inQFS = 3 And char = "(" Then
            parens = parens + 1
            fn = fn & char
        ElseIf inQFS = 3 And char = ")" Then
            parens = parens - 1
            fn = fn & char
            If parens = 0 Then
                Dim fnResolved
                fnResolved = ParseKeys(fn, cell, sheet, keys)
                If Not resolvable Or fnResolved = "" Then
                    resolvable = False
                ElseIf TypeName(fnResolved) = "Date" Then
                    fnResolved = CDbl(fnResolved)
                ElseIf TypeName(fnResolved) = "String" Then
                    fnResolved = """" & fnResolved & """"
                ElseIf TypeName(fnResolved) = "Boolean" And fnResolved Then
                    fnResolved = "TRUE"
                ElseIf TypeName(fnResolved) = "Boolean" And Not fnResolved Then
                    fnResolved = "FALSE"
                End If
                
                ' If resolvable, replace fn in resolved with fnResolved
                If resolvable Then
                    resolved = VBA.Left(resolved, VBA.Len(resolved) - VBA.Len(fn)) & fnResolved
                End If
                
                fn = ""
                inQFS = 0
                hasQFS = True
            End If
        ElseIf inQFS = 3 And parens > 0 Then
            fn = fn & char
        Else
            fn = ""
            inQFS = 0
        End If
    Next i
    
    ' If resolvable, return resolved
    If hasQFS And resolvable Then
        ParseFormula = sheet.Evaluate("=" & resolved)
    Else
        ParseFormula = ""
    End If
End Function

' Determine all finql keys required by a QFS formula
' Assumes that formula is a QFS formula (may have nested arguments)
' If formula can be resolved to a static value, this value is returned
Function ParseKeys(Formula As String, cell As Range, sheet As Worksheet, ByRef keys)
    ParseKeys = ""
    Dim argIndex As String: argIndex = VBA.InStr(Formula, "(")
    If argIndex = 0 Then Exit Function
    
    Dim name As String: name = VBA.UCase(VBA.Left(Formula, argIndex - 1))
    Dim args() As String: args = ParseArguments(Formula)
    Dim argsCount As Long: argsCount = NumElements(args)

    If name = "QFS" Or name = "=QFS" Or name = "=-QFS" Then
        Dim ticker As String
        Dim metric As String
        Dim period: period = ""
        Dim field As String
        Dim resolved As String
        Dim nested As Boolean
        
        ' Test each argument for nested QFS formulas
        ' and parse only the nested formulas since
        ' these must be resolved before we can determine
        ' the key for the current formula
        
        nested = False
        If argsCount > 0 Then
            If VBA.InStr(VBA.UCase(args(0)), "QFS(") > 0 Then
                ticker = ParseFormula(args(0), cell, sheet, keys)
                If ticker = "" Then nested = True
            End If
        End If
        
        If argsCount > 1 Then
            If VBA.InStr(VBA.UCase(args(1)), "QFS(") > 0 Then
                metric = ParseFormula(args(1), cell, sheet, keys)
                If metric = "" Then nested = True
            End If
        End If
        
        If argsCount > 2 Then
            If VBA.InStr(VBA.UCase(args(2)), "QFS(") > 0 Then
                period = ParseFormula(args(2), cell, sheet, keys)
                If period = "" Then nested = True
            End If
        End If
        
        ' This is currently not used. Include in preparation
        ' for additional arguments like 'min/max' for sector stats
        If argsCount > 3 Then
            If VBA.InStr(VBA.UCase(args(3)), "QFS(") > 0 Then
                field = ParseFormula(args(3), cell, sheet, keys)
                If field = "" Then nested = True
            End If
        End If
        
        If nested Then Exit Function
        
        ' Build the finql key required by the formula.
        ' This code is (sort of) duplicated in QFS, so
        ' if you change this, check that function as well
        
        If ticker = "" Then ticker = EvalArgument(args(0), cell, sheet)
        If metric = "" Then metric = EvalArgument(args(1), cell, sheet)
        
        ' Build finql key from arguments
        Dim index As Integer
        If argsCount > 2 Then
            If period = "" Then period = EvalArgument(args(2), cell, sheet)
            If TypeName(period) = "Double" And period < (Now() - (365 * 50)) Then
                index = CInt(period)
                period = ""
            ElseIf TypeName(period) = "Double" Or TypeName(period) = "Date" Then
                period = "Y" & Year(period) & ".M" & Month(period) & ".D" & Day(period)
            ElseIf TypeName(period) = "String" And IsDateString(CStr(period)) Then
                period = DateStringToPeriod(CStr(period))
            End If
        End If
        
        Dim key As String, withoutPeriod As String
        key = VBA.UCase(ticker) & "." & VBA.LCase(metric)
        withoutPeriod = key
        If period <> "" Then key = key & "[""" & VBA.UCase(period) & """]"

        ' If key is already cached, we can resolve this formula
        If IsCached(key) And Not IsCachedError(key) Then
            ParseKeys = CachedToQFS(key, index)
        End If

        ' Add resolved key to list of keys to request
        Call InsertElementIntoArray(keys, UBound(keys) + 1, key)
        
        ' Add the key without a period to the batch request because
        ' sometimes excel triggers the formula with a missing period
        ' argument (particularly on sheet load). Adding the key without
        ' its period to the batch prevents potentially unnecessary API
        ' requests when this happens, even though it may cost one
        ' extra datapoint.
        '
        ' NOTE:
        ' This scenario should now be resolved with an update to the
        ' QFS function that checks to make sure arguments match
        ' what is in the formula text. Thus, this is less-ideal solution
        ' is unused but left for reference in case more edge cases
        ' are discovered that are not covered by the QFS trick and
        ' we need to reactivate this optimization
        '
        ' Call InsertElementIntoArray(keys, UBound(keys) + 1, withoutPeriod)
        
    End If
End Function

' Parse a list of argument strings given an excel formula
Function ParseArguments(Formula As String) As String()
    Dim func As String
    Dim args() As String
    Dim safeArgs As String
    Dim c As String
    Dim i As Long, pdepth As Long
    Dim quoted As Boolean

    quoted = False
    func = VBA.Trim(Formula)
    i = VBA.InStr(func, "(")
    func = VBA.Mid(func, i + 1)
    func = VBA.Mid(func, 1, VBA.Len(func) - 1)

    ' Escape any commas in nested formulas or quotes
    For i = 1 To VBA.Len(func)
        c = VBA.Mid(func, i, 1)
        If c = "(" Or c = "[" Then
            pdepth = pdepth + 1
        ElseIf c = ")" Or c = "]" Then
            pdepth = pdepth - 1
        ElseIf c = """" Then
            quoted = Not quoted
        ElseIf (c = "," Or c = Application.International(xlListSeparator)) And pdepth = 0 And Not quoted Then
            c = "[[,]]"
        End If
        safeArgs = safeArgs & c
    Next i
    args = Split(safeArgs, "[[,]]")
    ParseArguments = args
End Function

' Evaluate the value of an argument that may include
' formulas or cell references
Function EvalArgument(arg As String, cell As Range, sheet As Worksheet)
    Dim value
    Dim address As String
    Dim resolvedCell As Range
    If IsCellAddress(arg) Then
        ' Evaluate reference to another sheet/cell
        Dim parts
        Dim sheetName As String
        Dim cellAddr As String
        
        parts = VBA.Split(arg, "!")
        If (NumElements(parts) > 1) Then
            sheetName = parts(0)
            If VBA.Left(sheetName, 1) = "'" Then sheetName = VBA.Mid(sheetName, 2, VBA.Len(sheetName) - 2)
            cellAddr = parts(1)
        Else
            sheetName = sheet.name
            cellAddr = parts(0)
        End If
            
        address = sheet.Parent.Sheets(sheetName).Range(cellAddr).address(External:=True)
        Set resolvedCell = Range(address)
        value = resolvedCell.value
        If IsEmpty(value) And resolvedCell.HasFormula Then
            value = resolvedCell.Worksheet.Evaluate(resolvedCell.Formula)
        End If

        EvalArgument = value
    ElseIf HasTableAddress(arg) Then
        ' Resolve table references
        value = sheet.Evaluate(ResolveTableAddresses(arg, cell))
        EvalArgument = value
    Else
        ' Evaluate nested formula arg or return constant value arg
        value = sheet.Evaluate(arg)
        EvalArgument = value
    End If
End Function

' Determine if string represents a valid excel cell address
Public Function IsCellAddress(strAddress As String) As Boolean
    Dim r As Range
    On Error Resume Next
    Set r = Range(strAddress)
    If Not r Is Nothing Then IsCellAddress = True
End Function

' Determine if argument includes implied-row
' references to a table cell
Function HasTableAddress(arg As String) As Boolean
    Dim i As Integer, inQuotes As Boolean, c As String
    
    HasTableAddress = False
    
    i = VBA.InStr(arg, "@")
    If i < 1 Then i = VBA.InStr(VBA.LCase(arg), "[#this row]")
    If i < 1 Then Exit Function
    
    inQuotes = False
    For i = 1 To VBA.Len(arg) - 1
        c = VBA.Mid(arg, i, 1)
        If c = """" Then
            ' Ignore quoted table references
            inQuotes = Not inQuotes
        ElseIf (Not inQuotes) And c = "[" Then
            Dim trimmed As String
            trimmed = VBA.LCase(VBA.Trim(VBA.Mid(arg, i + 1)))
            If VBA.Left(trimmed, 1) = "@" Or VBA.InStr(trimmed, "[#this row]") Then
                HasTableAddress = True
                i = VBA.Len(arg)
            End If
        End If
    Next i
End Function

' Replaces any table references with direct cell
' references relative to the given cell
Function ResolveTableAddresses(arg As String, cell As Range)
    Dim i As Integer, j As Integer, c As String, _
    inQuotes As Boolean, _
    inTable As String, _
    resolved As String, _
    tableName As String
    
    inQuotes = False
    inTable = ""
    
    ' Looking for the following syntax
    '   [ @ Header]
    '   [@  Header Name]
    '   [@ [Header Name]]
    '   [[#This Row] , Header]
    '   [[#this Row],[Header Name]]
    '   [[#this row],  Header Name]
    '   Table[@Header]
    '   Table[@[Header Name]]
    '   Table[[#This Row],Header]
    '   Table[[#This Row],[Header]]
    
    ' Assumptions:
    '  - No space allowed between table name and [
    '  - No brackets in header name
    '  - Extra space not allowed within row/col brackets
    '  - Escape characters are not recognized
    
    ' LogMessage "Resolving " & arg
    
    Dim validTableStart As String, validTableChar As String
    validTableStart = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz\._"
    validTableChar = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz._0123456789"
    
    For i = 1 To VBA.Len(arg)
        c = VBA.Mid(arg, i, 1)
        If c = """" Then
            inQuotes = Not inQuotes
            resolved = resolved & c
        ElseIf inQuotes Then
            resolved = resolved & c
        ElseIf inTable = "" And VBA.InStr(validTableStart, c) Then
            inTable = "name"
            tableName = c
        ElseIf inTable = "name" And VBA.InStr(validTableChar, c) Then
            tableName = tableName & c
        ElseIf (inTable = "" Or inTable = "name") And c = "[" Then
            Dim tableSpec As String
            inTable = "row"
            tableSpec = c
            For j = i + 1 To VBA.Len(arg)
                c = VBA.Mid(arg, j, 1)
                If c = " " Then
                    tableSpec = tableSpec & c
                ElseIf c = "@" Then
                    tableSpec = tableSpec & c
                    inTable = "col"
                    Exit For
                ElseIf c = "[" And VBA.LCase(VBA.Mid(arg, j, 11)) = "[#this row]" Then
                    tableSpec = tableSpec & VBA.Mid(arg, j, 11)
                    j = j + 10
                    inTable = "col"
                    Exit For
                Else
                    tableSpec = tableSpec & c
                    Exit For
                End If
            Next j
            i = j
            
            If inTable = "col" Then
                For j = i + 1 To VBA.Len(arg)
                    c = VBA.Mid(arg, j, 1)
                    If c = " " Or c = "," Or c = Application.International(xlListSeparator) Then
                        tableSpec = tableSpec & c
                    Else
                        Dim endPos As Integer, colName As String
                        endPos = VBA.InStr(j, arg, "]")
                        If endPos > 0 Then
                            colName = VBA.Mid(arg, j, endPos - j)
                            j = endPos
                            tableSpec = tableSpec & colName
                            ' If col name was bracketed, keep all space
                            ' inside brackets and consume to next closing bracket
                            If VBA.Left(colName, 1) = "[" Then
                                colName = VBA.Mid(colName, 2)
                                Dim k As Integer
                                For k = j To VBA.Len(arg)
                                    c = VBA.Mid(arg, k, 1)
                                    If c = " " Then
                                        tableSpec = tableSpec & c
                                    ElseIf c = "]" Then
                                        tableSpec = tableSpec & c
                                        Exit For
                                    Else
                                        Exit For
                                    End If
                                Next k
                                j = k
                            Else
                                ' Otherwise, trim extra space because it's
                                ' not part of column name
                                colName = VBA.Trim(colName)
                                j = j - 1
                            End If
                            
                            ' Parse to the end of the table spec
                            For k = j + 1 To VBA.Len(arg)
                                c = VBA.Mid(arg, k, 1)
                                If c = " " Then
                                    tableSpec = tableSpec & c
                                ElseIf c = "]" Then
                                    tableSpec = tableSpec & c
                                    Exit For
                                Else
                                    Exit For
                                End If
                            Next k
                            j = k
                        Else
                            inTable = ""
                            tableSpec = tableSpec & c
                        End If
                        Exit For
                    End If
                Next j
                i = j
            End If
            
            If inTable = "col" Then
                Dim table As ListObject
                If (tableName = "") Then
                    Set table = cell.ListObject
                Else
                    Set table = Application.Range(tableName & "[#All]").Parent.ListObjects(tableName)
                End If
                
                Dim row As Long, row2 As Long, addr As String
                row = table.DataBodyRange.Cells(1, 1).row - 1
                row2 = cell.row
                addr = table.DataBodyRange(row2 - row, table.ListColumns(colName).index).address(External:=True)
                resolved = resolved & addr
            Else
                resolved = resolved & tableName & tableSpec
            End If
            
            inTable = ""
            tableName = ""
        Else
            resolved = resolved & tableName & c
            inTable = ""
            tableName = ""
        End If
    Next i
    ResolveTableAddresses = resolved
    ' LogMessage "Resolved to " & resolved
End Function




---
title: VBIDE - Procedures - Line Numbers
# author: 'Anastasiou Alex'
# date: 2022-12-02 12:00:00 # if missing it's taken from the filename
# last_modified_at: #2022-12-02 12:00:00 
categories: [VBIDE, Procedures] # can handle 1 category and 1 subcategory eg [Category, Subcategory]
tags: [LineNumbers] # [Tag1, Tag2 ...]
---



```vb

Public Sub ModuleLinesNumbersAdd(Optional TargetWorkbook As Workbook, Optional Module As VBComponent)
'@AssignedModule F_Vbe_Lines_Number
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@INCLUDE PROCEDURE ProcedureLinesNumbersAdd
'@INCLUDE PROCEDURE ProceduresOfModule
'@INCLUDE PROCEDURE ActiveModule
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If Module Is Nothing Then Set Module = ActiveModule
    Dim Procedure
    For Each Procedure In ProceduresOfModule(Module)
        ProcedureLinesNumbersAdd TargetWorkbook, Module, CStr(Procedure)
    Next
End Sub

Public Sub ModuleLinesNumbersRemove( _
                                   Optional TargetWorkbook As Workbook, _
                                   Optional Module As VBComponent)
'@AssignedModule F_Vbe_Lines_Number
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@INCLUDE PROCEDURE ProcedureLinesNumbersRemove
'@INCLUDE PROCEDURE ProceduresOfModule
'@INCLUDE PROCEDURE ActiveModule
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If Module Is Nothing Then Set Module = ActiveModule
    Dim Procedure
    For Each Procedure In ProceduresOfModule(Module)
        ProcedureLinesNumbersRemove TargetWorkbook, Module, CStr(Procedure)
    Next
End Sub

Public Sub ProcedureLinesNumbersAdd( _
                                   Optional TargetWorkbook As Workbook, _
                                   Optional Module As VBComponent, _
                                   Optional Procedure As String, _
                                   Optional startNumberingAt = 1)
'@AssignedModule F_Vbe_Lines_Number
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ProcedureReplace
'@INCLUDE PROCEDURE IsLineNumberAble
'@INCLUDE PROCEDURE ProcedureCode
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Exit Sub
    Dim counter As Long
        counter = startNumberingAt
    Dim CodeLines
        CodeLines = Split(ProcedureCode(TargetWorkbook, Module, Procedure), vbNewLine)
    Dim Code As String
    Dim i As Long
    For i = LBound(CodeLines) To UBound(CodeLines)
        If Code = "" Then
            If IsLineNumberAble(CodeLines(i)) Then
                Code = counter & ":" & CodeLines(i)
                counter = counter + 1
            Else
                Code = CodeLines(i)
            End If
        Else
            If IsLineNumberAble(CodeLines(i)) And Right(Trim(CodeLines(i - 1)), 1) <> "_" Then
                Code = Code & vbNewLine & counter & ":" & CodeLines(i)
                counter = counter + 1
            Else
                Code = Code & vbNewLine & CodeLines(i)
            End If
        End If
    Next i
    ProcedureReplace Module, Procedure, Code
End Sub

Public Sub ProcedureLinesNumbersRemove( _
                                      Optional TargetWorkbook As Workbook, _
                                      Optional Module As VBComponent, _
                                      Optional Procedure As String)
'@AssignedModule F_Vbe_Lines_Number
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ProcedureReplace
'@INCLUDE PROCEDURE IndentCodeString
'@INCLUDE PROCEDURE ProcedureCode
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Exit Sub
    Dim startLine As Long
    Dim NumLines As Long
    Dim Code As String
    Dim CodeLines
        CodeLines = Split(ProcedureCode(TargetWorkbook, Module, Procedure), vbNewLine)
    Dim i As Long
    For i = LBound(CodeLines) To UBound(CodeLines)
        If Not IsNumeric(left(Trim(CodeLines(i)), 1)) Then
            Code = IIf(Code <> "", Code & vbNewLine, "") & CodeLines(i)
        Else
            CodeLines(i) = CodeLines(i) & " "
            Code = IIf(Code <> "", Code & vbNewLine, "") & _
                   Space(InStr(CodeLines(i), ":")) & _
                   Mid(CodeLines(i), InStr(CodeLines(i), ":") + 1)
        End If
    Next i
    Code = IndentCodeString(Code)
    ProcedureReplace Module, Procedure, Code
End Sub

Public Function IsLineNumberAble(ByVal str As String) As Boolean
'@AssignedModule F_Vbe_Lines_Number
    Dim Test As String
    Test = Trim(str)
    If Len(Test) = 0 Then Exit Function
    If Right(Test, 1) = ":" Then Exit Function
    If IsNumeric(left(Test, 1)) Then Exit Function
    If Test Like "'*" Then Exit Function
    If Test Like "Rem*" Then Exit Function
    If Test Like "Dim*" Then Exit Function
    If Test Like "Sub*" Then Exit Function
    If Test Like "Public*" Then Exit Function
    If Test Like "Private*" Then Exit Function
    If Test Like "Function*" Then Exit Function
    If Test Like "End Sub*" Then Exit Function
    If Test Like "End Function*" Then Exit Function
    If Test Like "Debug*" Then Exit Function
    IsLineNumberAble = True
End Function



```
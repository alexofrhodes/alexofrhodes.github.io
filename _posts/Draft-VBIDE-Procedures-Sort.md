---
title: VBIDE - Procedures - Sort
# author: 'Anastasiou Alex'
# date: 2022-12-02 12:00:00 # if missing it's taken from the filename
last_modified_at: #2022-12-02 12:00:00 
categories: [VBIDE, Procedures] # can handle 1 category and 1 subcategory eg [Category, Subcategory]
tags: [Sort] # [Tag1, Tag2 ...]
---


# Alphabetically

```vb

Public Sub ProceduresSortAlphabeticallyInModule(Optional Module As VBComponent)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE WorkbookOfModule
'@INCLUDE PROCEDURE ArrayQuickSort
'@INCLUDE PROCEDURE ProcedureCode
'@INCLUDE PROCEDURE ProceduresOfModuleArray
'@INCLUDE PROCEDURE ActiveModule

    If Module Is Nothing Then Set Module = ActiveModule
    If Module.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = WorkbookOfModule(Module)
    Dim Procedures
        Procedures = ProceduresOfModuleArray(Module)
    StartLine = Module.CodeModule.ProcStartLine(Procedures(0), vbext_pk_Proc)
    totalLines = Module.CodeModule.CountOfLines - Module.CodeModule.CountOfDeclarationLines
    ArrayQuickSort Procedures
    Dim ReplacedProcedures As String
    Dim i As Long
    For i = LBound(Procedures) To UBound(Procedures)
        ReplacedProcedures = IIf(ReplacedProcedures <> "", ReplacedProcedures & vbNewLine, "") & _
                             ProcedureCode(TargetWorkbook, Module, CStr(Procedures(i)))
    Next i
    Module.CodeModule.DeleteLines StartLine, totalLines
    Module.CodeModule.AddFromString ReplacedProcedures
End Sub

Public Sub ProceduresSortAlphabeticallyInWorkbook(Optional TargetWorkbook As Workbook)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@INCLUDE PROCEDURE ProceduresSortAlphabeticallyInModule
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        ProceduresSortAlphabeticallyInModule Module
    Next
End Sub

```

# Sub / Functions

```vb
Public Sub ProceduresSortSubFunctionInModule( _
                                            Optional TargetWorkbook As Workbook, _
                                            Optional Module As VBComponent)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE AssignWorkbookVariable
'@INCLUDE PROCEDURE AssignModuleVariable
'@INCLUDE PROCEDURE ArrayQuickSort
'@INCLUDE PROCEDURE ProcedureCode
'@INCLUDE PROCEDURE ProcedureType
'@INCLUDE PROCEDURE ProceduresOfModuleArray
'@INCLUDE PROCEDURE ModuleCodeRemove
    AssignWorkbookVariable TargetWorkbook
    AssignModuleVariable TargetWorkbook, Module
    If Module.CodeModule.CountOfLines = 0 Then Exit Sub

    Dim Procedures As Variant
        Procedures = ProceduresOfModuleArray(Module)
    ArrayQuickSort Procedures
    
    Dim TheSubs As String, TheFunctions As String, TheOther As String
    Dim code As String
    Dim Procedure As String
    Dim i As Long
    For i = LBound(Procedures) To UBound(Procedures)
        Procedure = CStr(Procedures(i))
        sProcedureType = ProcedureType(TargetWorkbook, Module, Procedure)
        code = ProcedureCode(TargetWorkbook, Module, Procedure)
        If sProcedureType = "Sub" Then
            TheSubs = IIf(TheSubs = "", code, TheSubs & vbNewLine & code)
        ElseIf sProcedureType = "Function" Then
            TheFunctions = IIf(TheFunctions = "", code, TheFunctions & vbNewLine & code)
        End If
    Next i
    ModuleCodeRemove Module
    'AddFromString Inserts the text starting on the line preceding the first procedure in the module.
    'If the module doesn't contain procedures, AddFromString places the text at the end of the module.
    Module.CodeModule.AddFromString TheFunctions
    Module.CodeModule.AddFromString TheSubs
End Sub

Public Sub ProceduresSortSubFunctionInWorkbook(Optional wb As Workbook)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@INCLUDE PROCEDURE ProceduresSortSubFunctionInModule
    If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule Then ProceduresSortSubFunctionInModule vbComp
    Next
End Sub

```

# Public / Private

```vb
Public Sub ProceduresSortPublicPrivateInModule(Optional Module As VBComponent)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE WorkbookOfModule
'@INCLUDE PROCEDURE ArrayQuickSort
'@INCLUDE PROCEDURE ProcedureCode
'@INCLUDE PROCEDURE ProcedureScope
'@INCLUDE PROCEDURE ProceduresOfModuleArray
'@INCLUDE PROCEDURE ModuleIgnore
'@INCLUDE PROCEDURE ActiveModule
'@INCLUDE PROCEDURE ModuleCodeRemove
    If Module Is Nothing Then Set Module = ActiveModule
    If ModuleIgnore(Module) Then Exit Sub
    If Module.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = WorkbookOfModule(Module)
    Dim ThePublic As String, ThePrivate As String, TheOther As String
    Dim code As String
    Dim Procedure As String
    Dim i As Long
    Dim Procedures
        Procedures = ProceduresOfModuleArray(Module)
    ArrayQuickSort Procedures
    For i = LBound(Procedures) To UBound(Procedures)
        Procedure = CStr(Procedures(i))
        code = ProcedureCode(TargetWorkbook, Module, Procedure)
        If ProcedureScope(Module, Procedure) = "Public" Then
            ThePublic = IIf(ThePublic = "", code, ThePublic & vbNewLine & code)
        Else
            ThePrivate = IIf(ThePrivate = "", code, ThePrivate & vbNewLine & code)
        End If
    Next i
    ModuleCodeRemove Module
    Module.CodeModule.AddFromString ThePrivate
    Module.CodeModule.AddFromString ThePublic
End Sub

Public Sub ProceduresSortPublicPrivateInWorkbook(Optional TargetWorkbook As Workbook)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@INCLUDE PROCEDURE ProceduresSortPublicPrivateInModule
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = vbext_ct_StdModule Then ProceduresSortPublicPrivateInModule Module
    Next
End Sub

```
# Helpers

```vb

Public Sub ArrayQuickSort(ByRef SortableArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
'@AssignedModule F_Arrays
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim varX As Variant
    If IsEmpty(SortableArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortableArray), "()") < 1 Then
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortableArray)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortableArray)
    End If
    If lngMin >= lngMax Then
        Exit Sub
    End If
    i = lngMin
    j = lngMax
    varMid = Empty
    varMid = SortableArray((lngMin + lngMax) \ 2)
    If IsObject(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If
    While i <= j
        While SortableArray(i) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortableArray(j) And j > lngMin
            j = j - 1
        Wend
        If i <= j Then
            varX = SortableArray(i)
            SortableArray(i) = SortableArray(j)
            SortableArray(j) = varX
            i = i + 1
            j = j - 1
        End If
    Wend
    If (lngMin < j) Then Call ArrayQuickSort(SortableArray, lngMin, j)
    If (i < lngMax) Then Call ArrayQuickSort(SortableArray, i, lngMax)
End Sub

```
---
title: VBIDE - Procedures - Reorder 
# author: 'Anastasiou Alex'
# date: 2022-12-02 12:00:00 # if missing it's taken from the filename
last_modified_at: #2022-12-02 12:00:00 
categories: [VBIDE, Procedures] # can handle 1 category and 1 subcategory eg [Category, Subcategory]
tags: [Sort] # [Tag1, Tag2 ...]
---


# Code

```vb

Sub ProcedureMoveDirectionUp( _
                            Optional TargetWorkbook As Workbook, _
                            Optional Module As VBComponent, _
                            Optional Procedure As String)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE CollectionIndexOf
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ProceduresOfModule
'@INCLUDE PROCEDURE ProcedureDelete
'@INCLUDE PROCEDURE ProcedureCode
'@INCLUDE PROCEDURE ProcedureTitleLineFirst
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    Dim Procedures As New Collection
    Set Procedures = ProceduresOfModule(Module)
    Dim index As Long
        index = CollectionIndexOf(Procedures, Procedure)
    If index = 1 Then Exit Sub
    Dim code As String
        code = ProcedureCode(TargetWorkbook, Module, Procedure)
    Dim PreviousProcedure As String
        PreviousProcedure = Procedures(index - 1)
    ProcedureDelete Module, Procedure
    Module.CodeModule.InsertLines ProcedureTitleLineFirst(Module, PreviousProcedure), code
    Dim ln As Long
    ln = ProcedureTitleLineFirst(Module, Procedure)
    Application.VBE.ActiveCodePane.SetSelection ln, 1, ln, 1
End Sub

Sub ProcedureMoveDirectionDown( _
                            Optional TargetWorkbook As Workbook, _
                            Optional Module As VBComponent, _
                            Optional Procedure As String)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE CollectionIndexOf
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ProceduresOfModule
'@INCLUDE PROCEDURE ProcedureLinesLast
'@INCLUDE PROCEDURE ProcedureDelete
'@INCLUDE PROCEDURE ProcedureCode
'@INCLUDE PROCEDURE ProcedureTitleLineFirst
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    Dim Procedures As New Collection
    Set Procedures = ProceduresOfModule(Module)
    Dim index As Long
        index = CollectionIndexOf(Procedures, Procedure)
    If index = Procedures.Count Then Exit Sub
    Dim code As String
        code = ProcedureCode(TargetWorkbook, Module, Procedure)
    Dim NextProcedure As String
        NextProcedure = Procedures(index + 1)
    ProcedureDelete Module, Procedure
    Module.CodeModule.InsertLines ProcedureLinesLast(Module, NextProcedure) + 1, code
    Dim ln As Long
    ln = ProcedureTitleLineFirst(Module, Procedure)
    Application.VBE.ActiveCodePane.SetSelection ln, 1, ln, 1
End Sub

Sub ProcedureMoveDirectionTop( _
                            Optional TargetWorkbook As Workbook, _
                            Optional Module As VBComponent, _
                            Optional Procedure As String)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE CollectionIndexOf
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ProceduresOfModule
'@INCLUDE PROCEDURE ProcedureDelete
'@INCLUDE PROCEDURE ProcedureCode
'@INCLUDE PROCEDURE ProcedureTitleLineFirst
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    Dim Procedures As New Collection
    Set Procedures = ProceduresOfModule(Module)
    Dim index As Long
        index = CollectionIndexOf(Procedures, Procedure)
    If index = 1 Then Exit Sub
    Dim code As String
        code = ProcedureCode(TargetWorkbook, Module, Procedure)
    Dim TopProcedure As String
        TopProcedure = Procedures(1)
    ProcedureDelete Module, Procedure
    Module.CodeModule.InsertLines ProcedureTitleLineFirst(Module, TopProcedure), code
    Dim ln As Long
    ln = ProcedureTitleLineFirst(Module, Procedure)
    Application.VBE.ActiveCodePane.SetSelection ln, 1, ln, 1
End Sub

Sub ProcedureMoveDirectionBottom( _
                            Optional TargetWorkbook As Workbook, _
                            Optional Module As VBComponent, _
                            Optional Procedure As String)
'@AssignedModule F_Vbe_Procedures
'@INCLUDE PROCEDURE CollectionIndexOf
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE PROCEDURE ProceduresOfModule
'@INCLUDE PROCEDURE ProcedureLinesLast
'@INCLUDE PROCEDURE ProcedureDelete
'@INCLUDE PROCEDURE ProcedureCode
'@INCLUDE PROCEDURE ProcedureTitleLineFirst
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop
    Dim Procedures As New Collection
    Set Procedures = ProceduresOfModule(Module)
    Dim index As Long
        index = CollectionIndexOf(Procedures, Procedure)
    If index = Procedures.Count Then Exit Sub
    Dim code As String
        code = ProcedureCode(TargetWorkbook, Module, Procedure)
    Dim LastProcedure As String
        LastProcedure = Procedures(Procedures.Count)
    ProcedureDelete Module, Procedure
    Module.CodeModule.InsertLines ProcedureLinesLast(Module, LastProcedure) + 1, code
    Dim ln As Long
    ln = ProcedureTitleLineFirst(Module, Procedure)
    Application.VBE.ActiveCodePane.SetSelection ln, 1, ln, 1
End Sub

```

# Helpers

```vb

Sub ProcedureDelete(Module As VBComponent, Procedure As String)
'@AssignedModule F_Vbe_Procedures
    Dim startLine As Long
    Dim NumLines As Long
    With Module.CodeModule
        startLine = .ProcStartLine(Procedure, vbext_pk_Proc)
        NumLines = .ProcCountLines(Procedure, vbext_pk_Proc)
        .DeleteLines startLine:=startLine, Count:=NumLines
    End With
End Sub

Public Function CollectionIndexOf(ByVal coll As Collection, _
                                  ByVal item As Variant, _
                                  Optional ByVal StartIndex As Long = 1) As Long
'@AssignedModule F_Collection
    Dim collindex As Long
    Dim collitemtype As Integer
    Dim itemtype As Integer
    
    itemtype = VarType(item)
    For collindex = StartIndex To coll.Count
        collitemtype = VarType(coll(collindex))
        If collitemtype = itemtype Then
            Select Case collitemtype
                Case 0 To 1: CollectionIndexOf = collindex: Exit Function
                Case 2 To 8, 11, 14, 17: If coll(collindex) = item Then CollectionIndexOf = collindex: Exit Function
                Case 9: If coll(collindex) Is item Then CollectionIndexOf = collindex: Exit Function
                Case Else
                    Debug.Print "Unsupported type for CollectionIndexOf."
                    Debug.Assert False
            End Select
        End If
    Next
    CollectionIndexOf = 0

End Function
```
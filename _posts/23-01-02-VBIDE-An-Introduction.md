---
title: Code that Reads and Modifies ... Code
# author: 'Anastasiou Alex'
# date: 2022-12-02 12:00:00 # if missing it's taken from the filename
last_modified_at: 2022-12-31 12:00:00 
categories: [VBIDE] # [Category, Subcategory]
# tags:  # [Tag1, Tag2 ...]
excerpt_separator: <!--more-->
---

<!--more-->

# Intro

With the VBA Extensibility Library (VBIDE) it is possible to 
write code that reads or modifies VBA projects, Modules, or Procedures. 
In this post I'll introduce some basic concepts and macros to work with the VBIDE.


## Getting started

**Add a Reference**  

  >In the VBEditor go to (Tools > References)  
  Find and tick "Microsoft Visual Basic For Applications Extensibility 5.3" 

**Enable Programmatic Access to the VBA Project**  

  >Switch to Excel, Options > Security > Macros > Trusted Publishers   
  Find and tick "Trust access to the Visual Basic Project"

**Project must be Unlocked**   

  >Your VB Project that you want to manipulate cannot be protected or locked.  

**Debugging**   

  >The code can be run but not stepped through,  if it makes changes to the project (eg. InsertLine)  

## Safety Practices

Some modifications can not be reverted. In case an unwanted change occurred and you cannot revert it with Undo, 
the only thing to do is to close the file without saving. **I can't stress enough the need to keep backups!**  

> Many VBA-based computer viruses propagate themselves by creating and/or modifying VBA code.  
Therefore, many virus scanners may automatically and without warning or confirmation 
delete modules that reference the VBProject object, 
causing a permanent and irretrievable loss of code.  
Consult the documentation for your anti-virus software for details. No liability is held for any loss or damage.

* **Do not open files from untrustworthy sources. In case of doubt, consider the file infected**
* **When possible, build it yourself (add the code/components manually)**
* **Read and understand the code before using it**
* **Use tools like [oleVba](https://www.decalage.info/python/olevba) to extract and analyze vba code without opening the file**

<hr>

# Content

## Macros - Workbook and Project

```vb 

{% include_relative CodeLibrary/procedures/ActiveCodepaneWorkbook.txt %}
{% include_relative CodeLibrary/procedures/WorkbookCode.txt %}
{% include_relative CodeLibrary/procedures/WorkbookExists.txt %}
{% include_relative CodeLibrary/procedures/WorkbookOfModule.txt %}
{% include_relative CodeLibrary/procedures/WorkbookProjectExists.txt %}
{% include_relative CodeLibrary/procedures/WorkbookProjectProtected.txt %}

```

## Module

A module is essentially a text file integrated into the parent application.  
Usually it is best to keep code related to similar tasks in the same module.   
Modules can be exported and imported.

**Code Modules**

The code modules are the most common place we store macros.  
The modules are located in the Modules folder within the workbook.

**Sheet Modules**

Each sheet in the workbook has a sheet object in the Microsoft Excel Objects folder.  
In its CodeModule where we can add event procedures (macros).  
These macros run when the user takes a specific action in the sheet.

**ThisWorkbook Module**

Each workbook contains one ThisWorkbook object at the bottom of the Microsoft Excel Objects folder.  
We can event based macros that run when the user takes actions in/on the workbook/sheets.

**Userforms**

Userforms are interactive windows where we can add controls like drop-downs, list boxes, check boxes etc.  
Each userform is stored in the Forms folder and has a **Designer** where we can add and edit controls before the Userform is opened 
and a **CodeModule** where we can put macros that will run when the form is open and/or the user interacts with the controls on the form.

**Class Modules**

Classes are stored in the Class Modules folder and allow us to write macros 
to create objects, properties, and methods. 
They can be used to create custom objects or collections that don't exist in the Object Library.

[source: excelcampus](https://www.excelcampus.com/vba/code-modules-event-procedures/)

<hr>

## Macros - Module

```vb

{% include_relative CodeLibrary/procedures/ActiveModule.txt %}
{% include_relative CodeLibrary/procedures/ModuleCode.txt %}
{% include_relative CodeLibrary/procedures/ModuleCodeContains.txt %}
{% include_relative CodeLibrary/procedures/ModuleExists.txt %}
{% include_relative CodeLibrary/procedures/ModuleExtension.txt %}
{% include_relative CodeLibrary/procedures/ModuleHeader.txt %}
{% include_relative CodeLibrary/procedures/ModuleHeaderContains.txt %}
{% include_relative CodeLibrary/procedures/ModuleOfProcedure.txt %}
{% include_relative CodeLibrary/procedures/ModuleTypeToString.txt %}

```

<hr>

## CodeModule

This represents the code behind a component, such as a form, class, or document.  
You use the CodeModule object to modify (add, delete, or edit) the code associated with a component and return information about the code text on a line-by-line basis. 
CodeModule methods which may come in handy:

|Name|Info|
|-|-|
|AddFromFile            |Inserts the contents of the file starting on the line preceding the first procedure in the code module. If the module doesn't contain procedures, AddFromFile places the contents of the file at the end of the module.|
|AddFromString          |Inserts the text starting on the line preceding the first procedure in the module. If the module doesn't contain procedures, AddFromString places the text at the end of the module.|
|CodePane               |Inserts a line or lines of code at a specified location in a block of code.|
|CountOfDeclarationLines|Returns a Long value indicating the number of lines of code in the Declarations section in a module.|
|CountOfLines           |Returns a Long value indicating the number of lines of code in a module.|
|CreateEventProc        |Create an event procedure and  returns the line at which the body of the event procedure starts.|
|DeleteLines            |Deletes a single line or a specified range of lines.|
|Find                   |Searches the active module for a specified string and returns True if a match is found.|
|InsertLines            |Inserts a line or lines of code at a specified location in a block of code.|
|Lines                  |Returns a string containing the contents of a specified line or lines.|
|Parent                 |The Module of the Codemodule.|
|ProcBodyLine           |Returns the number of the line at which the body of a specified procedure begins.|
|ProcCountLines         |Returns the number of lines in a specified procedure.|
|ProcOfLine             |Returns the name of the procedure that contains a specified line.|
|ProcStartLine          |Returns a value identifying the line at which a specified procedure begins.|
|ReplaceLine            |Replaces an existing line of code with a specified line of code.|

<hr>

## Codepane

The CodePane is a Window, the visible representation of the CodeModule.  
CodePane methods which may come in handy:

|Name|Info|
|-|-|
|GetSelection       |Returns the selection in a code pane.<br>Syntax:<br>object.GetSelection(startline, startcol, endline, endcol)|
|SetSelection       |Sets the selection in the code pane.<br>Syntax:<br>object.SetSelection(startline, startcol, endline, endcol)|
|TopLine            |Returns a Long specifying the line number of the line at the top of the code pane or sets the line showing at the top of the code pane. Read/write.|
|Window             |The window object|

## Macros - Codepane Selection

```vb

{% include_relative CodeLibrary/procedures/CodepaneSelection.txt %}
{% include_relative CodeLibrary/procedures/cpsColumnFirst.txt %}
{% include_relative CodeLibrary/procedures/cpsColumnLast.txt %}
{% include_relative CodeLibrary/procedures/cpsLineFirst.txt %}
{% include_relative CodeLibrary/procedures/cpsLineLast.txt %}
{% include_relative CodeLibrary/procedures/cpsLinesCode.txt %}
{% include_relative CodeLibrary/procedures/cpsLinesCount.txt %}
{% include_relative CodeLibrary/procedures/cpsPartAfter.txt %}
{% include_relative CodeLibrary/procedures/cpsPartBefore.txt %}
{% include_relative CodeLibrary/procedures/cpsSet.txt %}

```

<hr>

## Procedure

A set of commands to perform a specific task is placed into a procedure, 
which can be either a Subroutine or a Function.  
The main difference is that a Function procedure returns a result, whereas a Sub procedure does not.  
The return values have the following rules:
- The data type of the returned value must be declared in the Function header.
- The value to be returned must be assigned to a variable having the same name as the Function. 
- This variable does not need to be declared, as it already exists as a part of the function.

### Arguments

It is possible to pass data to a procedure via arguments, 
which are declared in the procedure definition.  

For example:

```vb
Sub AddToCells(i As Integer)
      .
      .
      .
End Sub
```

You can use multiple Optional arguments in a VBA procedure, 
as long the Optional arguments are all positioned at the end of the argument list.  
If they are omitted, the procedure will assign a default value to them.

For example:

```vb 
Sub AddToCells(Cell as Range, Optional i As Integer = 1)
```

### Passing Arguments By Value and By Reference

When arguments are passed to VBA procedures, they can be passed in two ways:

**ByVal**
Any changes that are made to the argument inside the procedure will be lost when the procedure ends.

```vb
Sub AddToCells(ByVal i As Integer)
        .
        .
        .
End Sub
```

**ByRef** 
Any changes that are made to the argument inside the procedure will be remembered when the procedure ends.  

```vb
Sub AddToCells(ByRef i As Integer)
        .
        .
        .
End Sub
```

***By default, VBA arguments are passed by Reference.***

[source: excelfunctions](https://www.excelfunctions.net/vba-functions-and-subroutines.html)


<hr>

## Macros - Procedure

<!-- ![Desktop View](/assets/img/ProcedureParts.png){: w="700" h="400" } -->

```vb
{% include_relative CodeLibrary/procedures/ActiveProcedure.txt %}
{% include_relative CodeLibrary/procedures/ProcedureBody.txt %}
{% include_relative CodeLibrary/procedures/ProcedureBodyLineFirst.txt %}
{% include_relative CodeLibrary/procedures/ProcedureBodyLineFirstAfterComments.txt %}
{% include_relative CodeLibrary/procedures/ProcedureBodyLineLast.txt %}
{% include_relative CodeLibrary/procedures/ProcedureCode.txt %}
{% include_relative CodeLibrary/procedures/ProcedureCodeContains.txt %}
{% include_relative CodeLibrary/procedures/ProcedureExists.txt %}
{% include_relative CodeLibrary/procedures/ProcedureHeader.txt %}
{% include_relative CodeLibrary/procedures/ProcedureHeaderContains.txt %}
{% include_relative CodeLibrary/procedures/ProcedureHeaderLineCount.txt %}
{% include_relative CodeLibrary/procedures/ProcedureHeaderLineFirst.txt %}
{% include_relative CodeLibrary/procedures/ProcedureHeaderLineLast.txt %}
{% include_relative CodeLibrary/procedures/ProcedureIgnore.txt %}
{% include_relative CodeLibrary/procedures/ProcedureKind.txt %}
{% include_relative CodeLibrary/procedures/ProcedureLinesCount.txt %}
{% include_relative CodeLibrary/procedures/ProcedureLinesFirst.txt %}
{% include_relative CodeLibrary/procedures/ProcedureLinesLast.txt %}
{% include_relative CodeLibrary/procedures/ProcedureReturnType.txt %}
{% include_relative CodeLibrary/procedures/ProcedureScope.txt %}
{% include_relative CodeLibrary/procedures/ProcedureTitle.txt %}
{% include_relative CodeLibrary/procedures/ProcedureTitleClean.txt %}
{% include_relative CodeLibrary/procedures/ProcedureTitleLineCount.txt %}
{% include_relative CodeLibrary/procedures/ProcedureTitleLineFirst.txt %}
{% include_relative CodeLibrary/procedures/ProcedureTitleLineLast.txt %}
{% include_relative CodeLibrary/procedures/ProcedureType.txt %}
{% include_relative CodeLibrary/procedures/ProceduresLike.txt %}
{% include_relative CodeLibrary/procedures/ProceduresOfModule.txt %}
{% include_relative CodeLibrary/procedures/ProceduresOfWorkbook.txt %}
```

<hr>

## Template

```vb


Sub VbeFormatTemplate( _
                     Optional TargetWorkbook As Workbook, _
                     Optional Module As VBComponent, _
                     Optional Procedure As String)

'@AssignedModule F_VbeFormat
'@INCLUDE PROCEDURE AssignCPSvariables

    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then Stop 'exit sub 'goto ErrorHandler
    'formatting code here
End Sub

Function AssignCPSvariables( _
                            ByRef TargetWorkbook As Workbook, _
                            ByRef Module As VBComponent, _
                            ByRef Procedure As String) As Boolean
    '

'@INCLUDE PROCEDURE AssignWorkbookVariable
'@INCLUDE PROCEDURE AssignProcedureVariable
'@INCLUDE PROCEDURE AssignModuleVariable
'@AssignedModule F_VbeFormat

    If Not AssignWorkbookVariable(TargetWorkbook) Then Exit Function
    If Not AssignProcedureVariable(TargetWorkbook, Procedure) Then Exit Function
    If Not AssignModuleVariable(TargetWorkbook, Module, Procedure) Then Exit Function
    AssignCPSvariables = True
    
End Function

Function AssignWorkbookVariable(ByRef TargetWorkbook As Workbook) As Boolean

'@INCLUDE PROCEDURE ActiveCodepaneWorkbook
'@AssignedModule F_VbeFormat
     If TargetWorkbook Is Nothing Then
        On Error Resume Next
        Set TargetWorkbook = ActiveCodepaneWorkbook
        On Error GoTo 0
    End If
    AssignWorkbookVariable = Not TargetWorkbook Is Nothing
End Function

Function AssignProcedureVariable(TargetWorkbook As Workbook, ByRef Procedure As String) As Boolean

'@INCLUDE PROCEDURE CodepaneSelection
'@INCLUDE PROCEDURE ActiveProcedure
'@INCLUDE PROCEDURE ProcedureExists
'@AssignedModule F_VbeFormat
    If Procedure = "" Then
        Dim cps As String
        cps = CodepaneSelection
        If Len(cps) > 0 Then
            Procedure = cps
        Else
            Procedure = ActiveProcedure
        End If
        If Not ProcedureExists(TargetWorkbook, Procedure) Then
            Debug.Print Procedure & " not found in Workbook " & TargetWorkbook.name
'            procedure = ""
        End If
    End If
    AssignProcedureVariable = Not Procedure = ""
End Function


Function AssignModuleVariable( _
                             ByVal TargetWorkbook As Workbook, _
                             ByRef Module As VBComponent, _
                             Optional ByVal Procedure As String) As Boolean

'@INCLUDE PROCEDURE CodepaneSelection
'@INCLUDE PROCEDURE ActiveModule
'@INCLUDE PROCEDURE ModuleOfProcedure
'@AssignedModule F_VbeFormat
    If Procedure = "" Then
        On Error Resume Next
        Set Module = ActiveModule
        On Error GoTo 0
    ElseIf Module Is Nothing Then
        On Error Resume Next
        Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
        On Error GoTo 0
    End If
    AssignModuleVariable = Not Module Is Nothing
End Function


```

<hr>

## Userform

Userforms have the **CodeModule**, which contains the code  
and the **Designer**, where we add and edit its permanent controls.  
  
Inside a userform, this code will place a commandbutton at runtime:

```vb  
dim ctr as MSForms.Control
Set ctr = me. Add(ProgID:="Forms.CommandButton.1",Name="cmb1", Visible:= True )  
ctr.Caption = "Click to Enter": ctr.top = 10: ctr.Left = 10  
```

We can add controls permanently with code like:

```vb
ActiveModule.designer.controls.Add "Forms.CommandButton.1","cmb1", True)
```
and we can work with the selected controls:

```vb
Public Function SelectedControl() As MSForms.control
'@INCLUDE PROCEDURE ActiveModue
    Dim Module As VBComponent
    Set Module = ActiveModule
    If SelectedControls.count = 1 Then
        Dim ctl    As control
        For Each ctl In Module.Designer.SELECTED
            Set SelectedControl = ctl
            Exit Function
        Next ctl
    End If
End Function

Public Function SelectedControls() As Collection
'@INCLUDE PROCEDURE ActiveModue
    Dim ctl    As control
    Dim out As New Collection
    Dim Module As VBComponent
    Set Module = ActiveModule
    For Each ctl In Module.Designer.SELECTED
        out.Add ctl
    Next ctl
    Set SelectedControls = out
    Set out = Nothing
End Function
```

Slight mod for selected controls inside Frames:

```vb
Public Function SelectedFrameControl() As MSForms.control
'@INCLUDE PROCEDURE ActiveModue
    Dim ctl    As control, c As control
    Dim out As New Collection
    Dim Module As VBComponent
    Set Module = ActiveModule
    
    For Each ctl In Module.Designer.SELECTED
        For Each c In ctl.Controls
            out.Add c
        Next
    Next ctl
    If out.count = 0 Then Exit Function
    Set SelectedFrameControl = out(1)
End Function

Public Function SelectedFrameControls() As Collection
'@INCLUDE PROCEDURE ActiveModue
    Dim ctl    As control, c As control
    Dim out As New Collection
    Dim Module As VBComponent
    Set Module = ActiveModule
    For Each ctl In Module.Designer.SELECTED
        For Each c In ctl.Controls
            out.Add c
        Next
    Next ctl
    Set SelectedFrameControls = out
    Set out = Nothing
End Function
```

<hr>

## Immediate Window

In the VBEditor learn to take advantage of the [Locals](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/locals-window) and [Watch](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/locals-window) Windows.  
That being said, the **Immediate Window** is your friend! Use it to:

* Test problematic or newly written code.
* Query or change the value of a variable while running an application. While execution is * halted, assign the variable a new value as you would in code.
* Query or change a property value while running an application.
* Call procedures as you would in code.
* View debugging output while the program is running.

Just type a line of code in the Immediate window and press ENTER to execute the statement.  
Note that they are executed as if they are entered in a specific module.  

If you need help on syntax for functions, statements, properties, or methods while working in the Immediate window, select the keyword, the property name, or the method name, and press F1.

We are not limited to simple strings either.

## Macros - Immediate Window

```vb

{% include_relative CodeLibrary/procedures/ArrayDimensions.txt %}
{% include_relative CodeLibrary/procedures/CalculateByteCharacters.txt %}
{% include_relative CodeLibrary/procedures/DPH.txt %}
{% include_relative CodeLibrary/procedures/DebugPrintHairetu.txt %}
{% include_relative CodeLibrary/procedures/DpHeader.txt %}
{% include_relative CodeLibrary/procedures/ShortenToByteCharacters.txt %}
{% include_relative CodeLibrary/procedures/TextDecomposition.txt %}
{% include_relative CodeLibrary/procedures/dp.txt %}
{% include_relative CodeLibrary/procedures/printArray.txt %}
{% include_relative CodeLibrary/procedures/printCollection.txt %}
{% include_relative CodeLibrary/procedures/printDictionary.txt %}
{% include_relative CodeLibrary/procedures/printRange.txt %}

```

<hr>

## Helpers


In this post the following were also used

```vb
{% include_relative CodeLibrary/procedures/ArrayRemoveEmptyElements.txt %}
{% include_relative CodeLibrary/procedures/ArrayTrim.txt %}
{% include_relative CodeLibrary/procedures/RemoveBlankLines.txt %}

```

<hr>

# Afterword
   
<!-- Enough for an introduction.   -->

<!-- There are lots of interesting things to discuss but they're all over the place.  
Hopefully this blog will help me review my code library and provide you with interesting content.   -->


<!-- YouTube and GitHub need to be revamped -->
<!-- @TODO create a file where to keep all the code -->
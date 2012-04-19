'******************************************************************************
'* File:     generate.vbs
'* Purpose:  Generate resource And script from model
'
'* Title:
'* Category:
'* Version:  1.0
'* Company:  Epicentre
'******************************************************************************

' Args[0]: Path of model file *.pdm
' Args[1]: Space

Set oArgs = WScript.Arguments

Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")

Include oFileSystemObject.getFolder(".") & "\utils.vbs"

StrModel = ""

If oArgs.count()>0 Then
  StrModel = oArgs(0)
Else
  Set oFolder = oFileSystemObject.getFolder(".").ParentFolder.ParentFolder
  Set oFiles = oFolder.Files
  If oFiles.Count <> 0 Then
    For Each oFile in oFolder.Files
      if lcase(oFileSystemObject.GetExtensionName(oFile.Name)) = "pdm" then
        StrModel = oFile.Path
        Exit For
      End If
    Next
  End If
End If

WScript.Echo StrModel

If Not IsObject(ActiveModel) Then
  Set oApp = CreateObject("PowerDesigner.Application")
      oApp.InteractiveMode = 0
      oApp.Locked = False
  Set oModel = oApp.OpenModel(StrModel, omf_DontOpenView Or omf_Hidden)
Else
  Set oModel = ActiveModel
  strPathModel = oFileSystemObject.GetParentFolderName(ActiveModel.Filename)
End If

strPathModel = CreateFolder(oFileSystemObject.GetParentFolderName(oModel.Filename))
strPathSQL = CreateFolder(strPathModel & "\sql")

Set opts = oModel.GetPackageOptions()

opts.GenerationCreateTrigger = false

' Create SQL Script for database creation
' .......................................
WScript.Echo "Begin script generation (" & oModel.DBMS.Code & ")"

set oTables = SortCollection(oModel.Tables)
set oViews = SortCollection(oModel.Views)
set oDomains = SortCollection(oModel.Domains)

Wiki = ""

WScript.Echo "Cr�ation des tables (fichiers)"

For Each oTable In oTables
   If IsObject(oTable) Then
      WScript.Echo "  " & oTable.Name
            
      CodeMax = 0
      FieldMax = 0
      
      For Ni=0 to oTable.Columns.Count -1
         Set oColumn=oTable.Columns.Item(Ni)
         If IsObject(oColumn) And Not (oColumn.Computed) Then
            If CodeMax < Len(oColumn.Code) then 
              CodeMax = Len(oColumn.Code)
            End if
            If FieldMax < Len(oColumn.Length) then 
              FieldMax = Len(oColumn.Length)
            End if
         End If
      Next

      Desc =        Space(CodeMax + 1) & "|" & String(40 + FieldMax, "=") & vbCrLf
	  Desc = Desc & Space(CodeMax + 1) & "|  SURVEY TITLE" & vbCrLf
	  Desc = Desc & Space(CodeMax + 1) & "|  FORM" & vbCrLf
	  Desc = Desc & Space(CodeMax + 1) & "|" & String(40 + FieldMax, "=") & vbCrLf
	  Desc = Desc & Space(CodeMax + 1) & "|"
	  
      For Ni=0 to oTable.Columns.Count -1
         Set oColumn=oTable.Columns.Item(Ni)
         If IsObject(oColumn) And Not (oColumn.Computed) Then
            If oColumn.Comment="" Then
              FieldLabel = oColumn.Name
            Else
              FieldLabel = oColumn.Comment
            End if
            Desc = Desc & oColumn.Code & Space(CodeMax + 1 -Len(oColumn.Code)) & "|  " & FieldLabel & Space(40-Len(FieldLabel))
            If Mid(oColumn.DataType, 1, 7) = "VARCHAR" Then
              Desc = Desc & "@<A" & Space(oColumn.Length) & ">"
            End If
            Desc = Desc & vbCrLf
         End If
      Next
   
	  Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\" & LCase(oTable.Code) & ".qes", ForWriting, true)
	  oFile.Write Desc & vbCrLf
	  oFile.Close
	  
   End If
   
Next

Set oApp = Nothing

WScript.Quit

' Functions
' .......................................

Function Include (Scriptname)
    Set oFile = oFileSystemObject.OpenTextFile(Scriptname)
    ExecuteGlobal oFile.ReadAll()
    oFile.Close
End Function

Function ExtendedAttribute (Column, AttributeName)
    s=Column.GetExtendedAttribute(oModel.DBMS.Code & "." & AttributeName)
    s=replace(s, vbCrLf, "")
    s=replace(s, "'", "\'")
    ExtendedAttribute = s
End Function
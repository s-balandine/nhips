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
  Set oFolder = oFileSystemObject.getFolder(".").ParentFolder
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

If StrModel="" Then
  StrModel = "model.pdm"
End If

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

WScript.Echo "Création des tables (fichiers)"

For Each oTable In oTables
   If IsObject(oTable) And (oTable.Name="Household") Then
   
      WScript.Echo "  " & oTable.Name
            
      Template = oTable.BeginScript
      Template = Replace(Template, vbCrLf, "\n")
      
      
      Set regEx = New RegExp
      regEx.IgnoreCase = True
      'regEx.Global = True
      
      Do
      
	      regEx.Pattern = "¤+"
	      Set Matches1 = regEx.Execute(Template)
	      regEx.Pattern = ":[A-Z_]+"
	      Set Matches2 = regEx.Execute(Template)
	
	      If Matches1.Count=0 Then Exit do
	      
	      Set Match1 = Matches1.Item(0)
	      Set Match2 = Matches2.Item(0)
	      
	      If Match2.Value =":HSH_IDENTIFIER_OF_CLUSTER" Then Exit do
	      
	      Set oColumn=Nothing
	      
	      For Each oColumn in oTable.Columns
	        If IsObject(oColumn) And Not (oColumn.Computed) And (oColumn.Code=Mid(Match2.Value, 2)) Then
	            Exit for
	         End If
	      Next 
	      
	      WScript.Echo "    " & Match1.Value & ":" & Match2.Value & vbCrLf
	      
	      If oColumn.DataType="SMALLINT" Then
	        Template = Mid(Template, 1, Match1.FirstIndex) + String(Match1.Length, "#") + Mid(Template, Match1.FirstIndex + Match1.Length)
	        Template = Mid(Template, 1, Match2.FirstIndex) +                              Mid(Template, Match2.FirstIndex + Match2.Length)
	      ElseIf oColumn.DataType="INTEGER" Then
	        Template = Mid(Template, 1, Match1.FirstIndex) + String(Match1.Length, "#") + Mid(Template, Match1.FirstIndex + Match1.Length)
	        Template = Mid(Template, 1, Match2.FirstIndex) +                              Mid(Template, Match2.FirstIndex + Match2.Length)
	      ElseIf oColumn.DataType="LONG" Then
	        Template = Mid(Template, 1, Match1.FirstIndex) + String(Match1.Length, "#") + Mid(Template, Match1.FirstIndex + Match1.Length)
	        Template = Mid(Template, 1, Match2.FirstIndex) +                              Mid(Template, Match2.FirstIndex + Match2.Length)
	      ElseIf oColumn.DataType="VARCHAR" Then
	        Template = _
	          Mid(Template, 1, Match1.FirstIndex) + "@<A" & Space(oColumn.Length) & ">" + _
	          Mid(Template, Match1.FirstIndex + Match1.Length)
	      End If
      
      Loop
      
      Template = Replace(Template, "\n", vbCrLf)
   
	  Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\" & LCase(oTable.Code) & ".qes", ForWriting, true)
	  oFile.Write Template & vbCrLf
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

Function RegExpTest(patrn, strng)
  Dim regEx, retVal            ' Create variable.
  Set regEx = New RegExp         ' Create regular expression.
  regEx.Pattern = patrn         ' Set pattern.
  regEx.IgnoreCase = False      ' Set case sensitivity.
  retVal = regEx.Test(strng)      ' Execute the search test.
  If retVal Then
    RegExpTest = "One or more matches were found."
  Else
    RegExpTest = "No match was found."
  End If
End Function

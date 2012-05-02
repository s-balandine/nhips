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
strPathSQL = CreateFolder(strPathModel & "\epidata")

Set opts = oModel.GetPackageOptions()

opts.GenerationCreateTrigger = false

' Create SQL Script for database creation
' .......................................
WScript.Echo "Begin script generation (" & oModel.DBMS.Code & ")"

set oTables = SortCollection(oModel.Tables)
set oViews = SortCollection(oModel.Views)
set oDomains = SortCollection(oModel.Domains)

Wiki = ""

WScript.Echo "Création des fichiers QES"

NCharWidth = 80
NCharMax = 80

Form="Eligible Woman"
'Form="Survey Supervisor"

For Each oTable In oTables
    
	If IsObject(oTable) And (oTable.Name=Form) Then
	'If IsObject(oTable) Then	
				
		WScript.Echo "  " & oTable.Name
		
		NCharMaxColumnName = 0
		NCharMaxColumnSize = 0
		
		SectionFirst = ExtendedAttribute (oTable, "SectionFirst")
		'SectionFirst = 1
		
		For Each oColumn in oTable.Columns
			If IsObject(oColumn) And Not (oColumn.Computed) Then

				ColumnSize = 0
			    ColumnName = ExtendedAttribute(oColumn, "Label")
			    
			    If Mid(oColumn.DataType, 1, 7)="NUMERIC" Then
					ColumnSize = oColumn.Length	
				End If
				
				If NCharMaxColumnName < Len(ColumnName) Then
					NCharMaxColumnName = Len(ColumnName)
				End If
				
				If NCharMaxColumnSize < oColumn.Length	 Then
					NCharMaxColumnSize = oColumn.Length	
				End If
				
			End If
		Next 

		If NCharMaxColumnName > NCharMaxColumnSize Then
			NCharMax = NCharMaxColumnName + 2
		Else
		    NCharMax = NCharMaxColumnSize + 2
		End If
		
		NCharMax = NCharMaxColumnSize + 2
		NCharMax = 34
		
		ColumnSectionN = 0
		ColumnSection = ""
		ColumnSectionPrev = ""
	    ColumnQuestionN = 1
	    ColumnQuestion = ""
	    ColumnQuestionPrev = ""
		ColumnNamePrev = ""
		ColumnN = 0
		
		Desc = String(NCharWidth, "=") & vbCrLf
		Desc = Desc & ExtendedAttribute (oModel, "Title") & vbCrLf
		Desc = Desc & ExtendedAttribute (oTable, "Title")

		For Each oColumn in oTable.Columns
			If IsObject(oColumn) And oColumn.Primary Then
			    WScript.Echo "    Key Primary: " & oColumn.Name
				If oColumn.DataType="AUTOINCREMENT" Then
					Desc = Desc & Space(NCharWidth - Len(ExtendedAttribute(oTable, "Title")) - 22)
					Desc = Desc & "{Rec}ord {ID}: <IDNUM>" & vbCrLf
				Else
				    Desc = Desc & Space(NCharWidth - Len(ExtendedAttribute(oTable, "Title")) - oColumn.Length - 15)
				    Desc = Desc & "{Rec}ord {ID}: <A" & Space(oColumn.Length) & ">" & vbCrLf
				End If
			End If
		Next

		Desc = Desc & String(NCharWidth, "=") & vbCrLf

		For Each oColumn in oTable.Columns
			If IsObject(oColumn) And Not (oColumn.Computed) And Not (oColumn.Primary) Then
				WScript.Echo "    Question: " & oColumn.Name
				
				ColumnName = ExtendedAttribute (oColumn, "Label")
				ColumnSection = ExtendedAttribute (oColumn, "Section")
				ColumnQuestion = ExtendedAttribute (oColumn, "Question")
				
				If ColumnSection<>ColumnSectionPrev Then
				    ColumnSectionN = ColumnSectionN + 1
				    ColumnSectionPrev = ColumnSection
					If ColumnSectionN - SectionFirst < 0 Then
						ColumnPrefix = "S"
						ColumnSectionNOffset = ColumnSectionN
					Else
						ColumnPrefix = "Q"
						ColumnSectionNOffset = ColumnSectionN - SectionFirst + 1
					End If
					Desc = Desc & vbCrLf
					Desc = Desc & ColumnSectionNOffset & "." & UCase(ColumnSection) & vbCrLf
					Desc = Desc & String(NCharWidth, "=") & vbCrLf
					ColumnQuestionN = 1
				End If
						
				If ColumnQuestion<>ColumnQuestionPrev Then
				    ColumnQuestionPrev = ColumnQuestion
				    Desc = Desc & String(NCharWidth, "-") & vbCrLf
					If ColumnQuestionN > 10 Then 
						Desc = Desc & ColumnPrefix & ColumnSectionNOffset & ColumnQuestionN & Space(2)
					Else
						Desc = Desc & ColumnPrefix & ColumnSectionNOffset & "0" & ColumnQuestionN & Space(2)
					End If
					If Len(ColumnQuestion) + 6 <= NCharWidth Then 
						Desc = Desc & ColumnQuestion
					Else
						ColumnQuestionPart1 = Mid(ColumnQuestion, 1, InStrRev(ColumnQuestion, " ", NCharWidth-6)) 
						ColumnQuestionPart2 = Mid(ColumnQuestion, InStrRev(ColumnQuestion, " ", NCharWidth-6)+1) 
						Desc = Desc & ColumnQuestionPart1 & vbCrLf & Space(6) & ColumnQuestionPart2
					End If
					If Len(ColumnQuestion) > (NCharWidth - NCharMax - 12 - 6) Then 
					    ColumnQuestionNotBreak = False
						Desc = Desc & vbCrLf & vbCrLf
					Else
					    ColumnQuestionNotBreak = True
					End If
				    ColumnQuestionN = ColumnQuestionN + 1
				    If ExtendedAttribute(oColumn, "CheckOffset") <> "" Then
						ColumnQuestionN = ColumnQuestionN + ExtendedAttribute(oColumn, "CheckOffset")
					End if
				    ColumnN = 0
				End If
				
				ColumnN = ColumnN + 1
								
				If Not oColumn.CannotModify Then
									
					If ColumnQuestionNotBreak Then
						If NCharWidth - NCharMax - 12 - Len(ColumnQuestion) - 6 > 0 Then 
					    	Desc = Desc & Space(NCharWidth - NCharMax - 12 - Len(ColumnQuestion) - 6 - 1)
					    End If
					    ColumnQuestionNotBreak = False
					Else				
					    Desc = Desc & Space(NCharWidth - NCharMax - 12 - 1)
					End If
					
					If ColumnQuestionN > 10 Then 
						ColumnNameEpiData = "({" & ColumnPrefix & ColumnSectionNOffset & ColumnQuestionN-1 & "." & ColumnN & "})" & Space(2)
						SetExtendedAttribute oColumn, "NameEpiData", ColumnPrefix & ColumnSectionNOffset & ColumnQuestionN-1 & ColumnN
					Else
						ColumnNameEpiData = "({" & ColumnPrefix &  ColumnSectionNOffset & "0" & ColumnQuestionN-1 & "." & ColumnN & "})" & Space(2)
						SetExtendedAttribute oColumn, "NameEpiData", ColumnPrefix & ColumnSectionNOffset & "0" & ColumnQuestionN-1 & "." & ColumnN
					End If
								
					If Mid(oColumn.DataType, 1, 7)="NUMERIC" Then
					    Desc = Desc & ColumnNameEpiData
					    If (Len(ColumnName)+ oColumn.Length + 2) >= NCharMax Then
							Desc = Desc & Mid(ColumnName, 1, NCharMax - oColumn.Length - 2) 
							Desc = Desc & String(2, ".")
						Else
							Desc = Desc & ColumnName
							Desc = Desc & String(NCharMax - Len(ColumnName) - oColumn.Length, ".")
						End If
						Desc = Desc & String(oColumn.Length, "#") & "  "
					End If
									
					If Mid(oColumn.DataType, 1, 7)="VARCHAR" Then
						If (Len(ColumnName)+ oColumn.Length + 6) >= NCharMax Then
							Desc = Desc & Space(12) & ColumnName & ":" & vbCrLf
							Desc = Desc & Space(NCharWidth - NCharMax - 13)
							Desc = Desc & ColumnNameEpiData & Space(NCharMax - oColumn.Length +1)
						Else
							Desc = Desc & ColumnNameEpiData
							Desc = Desc & ColumnName
							Desc = Desc & String(NCharMax - Len(ColumnName) - oColumn.Length, ".")
						End If
						Desc = Desc & "<A" & String(oColumn.Length - 1, " ") & ">"
					End If
									
					Desc = Desc & vbCrLf
					
				End If
			End If
		Next 
		
		Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\" & LCase(oTable.Name) & ".qes", ForWriting, true)
		oFile.Write Desc & vbCrLf
		oFile.Close
	End If
Next

WScript.Echo "Création des fichiers CHK"

For Each oTable In oTables
	If IsObject(oTable) And (oTable.Name=Form) Then
	'If IsObject(oTable) Then		
		WScript.Echo "  " & oTable.Name
				
		Desc = "LABELBLOCK" & vbCrLf
		
		For Each oDomain In oDomains
			If IsObject(oDomain) And (oDomain.ListOfValues<>"") Then
				WScript.Echo "    Labels: " & oDomain.Name
				Values = oDomain.ListOfValues
				Values = Split(Values, vbNewLine, -1, 1)
				Desc = Desc & "  LABEL " & UCase(oDomain.Code) & vbCrLf
				For i=0 To UBound(Values)
					If Values(i) <> "" Then
						Value = Values(i)
						Value = Split(Value, vbTab, -1, 1)
						If InStr(Value(1), "-") > 0 Then
							Desc=Desc & "    " & Value(0) & " """ & Mid(Value(1), InStr(Value(1), "-") + 1) & """" & vbCrLf 
						Else
						    Desc=Desc & "    " & Value(0) & " """ & Value(1) & """" & vbCrLf 
						End If
					End If
				Next
				Desc = Desc & "   END" & vbCrLf
			End If
		Next
		
		Desc = Desc & "END" & vbCrLf & vbCrLf

		Desc = Desc & "BEFORE RECORD" & vbCrLf
		For Each oColumn in oTable.Columns
		    If IsObject(oColumn) And (oColumn.Mandatory) And (oColumn.DefaultValue<>"") And Not (oColumn.CannotModify) Then
			    WScript.Echo "    DefaultValue: " & oColumn.Name
			    ColumnName = UCase(Replace(ExtendedAttribute(oColumn, "NameEpiData"), ".", ""))
				Desc = Desc & " IF (" & ColumnName & " = .) THEN" & vbCrLf
				Desc = Desc & "  LET " & ColumnName & "=" & oColumn.DefaultValue & vbCrLf
				Desc = Desc & " ENDIF" & vbCrLf
			End If
		Next
		Desc = Desc & "END" & vbCrLf & vbCrLf

		Desc = Desc & "AFTER RECORD" & vbCrLf
		For Each oColumn in oTable.Columns
		    If IsObject(oColumn) And (oColumn.Mandatory) And Not (oColumn.CannotModify) Then
			    WScript.Echo "    Mandatory: " & oColumn.Name
			    ColumnName = UCase(Replace(ExtendedAttribute(oColumn, "NameEpiData"), ".", ""))
			    ColumnLabel = ExtendedAttribute (oColumn, "Label")
				Desc = Desc & " IF (" & ColumnName & " = .) THEN" & vbCrLf
				Desc = Desc & "  HELP """ & ColumnLabel & " must be entered"" TYPE=ERROR" & vbCrLf
				Desc = Desc & "  GOTO " & ColumnName & "" & vbCrLf
				Desc = Desc & "  EXIT" & "" & vbCrLf
				Desc = Desc & " ENDIF" & vbCrLf
			End If
		Next
		Desc = Desc & "END" & vbCrLf


		For Each oColumn in oTable.Columns
			If IsObject(oColumn) And Not (oColumn.Computed) Then
				ColumnName = oColumn.Name
				ColumnCode = oColumn.Code
				
				WScript.Echo "    Attribute: " & ColumnName & " (" & ColumnCode & ")"
				ColumnName = UCase(Replace(ExtendedAttribute(oColumn, "NameEpiData"), ".", ""))
				Desc = Desc & vbCrLf
				Desc = Desc & "* " & ExtendedAttribute(oColumn, "NameEpiData") & " | "
				Desc = Desc & ColumnName & " (" & ColumnCode & ")" & " | "
				Desc = Desc & oColumn.DataType & vbCrLf
				Desc = Desc & ExtendedAttribute(oColumn, "Question")
				If ExtendedAttribute(oColumn, "Label") <> "" Then 
					Desc = Desc & " > " & ExtendedAttribute(oColumn, "Label") & vbCrLf
				Else
				    Desc = Desc & vbCrLf
				End If
				If oColumn.Domain.ListOfValues <> "" Then
					Values = oColumn.Domain.ListOfValues
					Values = Split(Values, vbNewLine, -1, 1)
					Desc = Desc & "* Allowed Values" & vbCrLf
					For i=0 To UBound(Values)
						If Values(i) <> "" Then
							Value = Values(i)
							Value = Split(Value, vbTab, -1, 1)
							Desc=Desc & "*    " & Value(0) & ": " & Value(1) & vbCrLf 
						End If
					Next
				End If
				If oColumn.Primary Then 
				    Desc = Desc & "RECID" & vbCrLf		
				    Desc = Desc & "  KEY UNIQUE 1" & vbCrLf 
					Desc = Desc & "  NOENTER" & vbCrLf 
				ElseIf Not (oColumn.CannotModify) Then
					ColumnName = UCase(Replace(ExtendedAttribute(oColumn, "NameEpiData"), ".", ""))
					Desc = Desc & ColumnName & vbCrLf		
					If oColumn.Mandatory Then Desc = Desc & "  MUSTENTER" & vbCrLf
					If oColumn.LowValue<>"" And oColumn.HighValue<>"" Then 
						Desc = Desc & "  RANGE " & oColumn.LowValue & " " & oColumn.HighValue & vbCrLf 
						If oColumn.Domain.ListOfValues <> "" Then
							Values = oColumn.Domain.ListOfValues
							Values = Split(Values, vbNewLine, -1, 1)
							Desc = Desc & "  MISSINGVALUE "
							For i=0 To UBound(Values)
								If Values(i) <> "" Then
									Value = Values(i)
									Value = Split(Value, vbTab, -1, 1)
									Desc=Desc & " " & Value(0)
								End If
							Next	
							Desc = Desc & vbCrLf
						End If
					Else
						If oColumn.Domain.ListOfValues <> "" Then
							Desc = Desc & "  COMMENT LEGAL USE " & UCase(oColumn.Domain.Code) & " SHOW" & vbCrLf
							Desc = Desc & "  TYPE COMMENT" & vbCrLf
						End If
					End If
					S = ExtendedAttribute(oColumn, "Check")
					If Len(S) > 0 Then
						S = Mid(S, 1, Len(S)-3)
						Desc = Desc & "  " & Replace(S, "  ", vbCrLf & "    ") & vbCrLf & "  END" & vbCrLf 
						For Each oBusinessRule in oColumn.AttachedRules
							If IsObject(oBusinessRule) Then
								'Desc = Desc & "  " & Replace(oBusinessRule.ServerExpression, "::", "  " & vbCrLf) & vbCrLf 
							End if
						Next
					End If
				End If
				Desc = Desc & "END" & vbCrLf	
			End If
		Next 
		
		For Each oColumn in oTable.Columns
	       If IsObject(oColumn) Then
				ColumnName = UCase(Replace(ExtendedAttribute(oColumn, "NameEpiData"), ".", ""))
				Desc = Replace(Desc, oColumn.Code & " ", ColumnName & " ")
				Desc = Replace(Desc, oColumn.Code & vbCrLf, ColumnName & vbCrLf) 
			End If
		Next
		
		Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\" & LCase(oTable.Name) & ".chk", ForWriting, true)
		oFile.Write Desc & vbCrLf
		oFile.Close
		
		
	End If
Next


Set oApp = Nothing

WScript.Quit

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

Function ExtendedAttribute (Object, AttributeName)
    s=Object.GetExtendedAttribute(oModel.DBMS.Code & "." & AttributeName)
    s=replace(s, vbCrLf, "")
    s=replace(s, "'", "\'")
    ExtendedAttribute = s
End Function

Function SetExtendedAttribute (Object, AttributeName, Value)
    Object.SetExtendedAttribute oModel.DBMS.Code & "." & AttributeName, Value 
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

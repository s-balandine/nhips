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

Form="Office Keyer"
FormAll= False

For Each oTable In oTables
    
	If IsObject(oTable) And (FormAll Or (oTable.Name=Form)) Then	
				
		WScript.Echo "  " & oTable.Name
		
		NCharMaxColumnName = 0
		NCharMaxColumnSize = 0
		
		SectionFirst = ExtendedAttribute (oTable, "SectionFirst")
		
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
				Desc = Desc & Space(NCharWidth - Len(ExtendedAttribute(oTable, "Title")) - oColumn.Length - 3)
				Desc = Desc & "<A" & Space(oColumn.Length) & ">" & vbCrLf
			End If
		Next
						
		Desc = Desc & String(NCharWidth, "=") & vbCrLf

		For Each oColumn in oTable.Columns
			If IsObject(oColumn) And Not (oColumn.Computed) And Not (oColumn.Primary) And (oColumn.Name<>"Identifier (Natural)") Then
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
					If ColumnSectionN=1 Then
						ColumnSection = "Identification"
					End If 
					Desc = Desc & ColumnSectionNOffset & "." & UCase(ColumnSection) & vbCrLf
					Desc = Desc & String(NCharWidth, "=") & vbCrLf
					ColumnQuestionN = 1
				End If
						
				If ColumnQuestion<>ColumnQuestionPrev Then
				    ColumnQuestionPrev = ColumnQuestion
				    If ColumnQuestionN > 0 Then 
				    	Desc = Desc & String(NCharWidth, "-") & vbCrLf
				    End If
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
					
					If oColumn.ForeignKey Then
						ColumnNameEpiData = "({" & ExtendedAttribute(oColumn, "NameEpiData") & "})   " & Space(2)
					Else
						If ColumnQuestionN > 10 Then 
							ColumnNameEpiData = "({" & ColumnPrefix & ColumnSectionNOffset & ColumnQuestionN-1 & "." & ColumnN & "})" & Space(2)
							SetExtendedAttribute oColumn, "NameEpiData", ColumnPrefix & ColumnSectionNOffset & ColumnQuestionN-1 & ColumnN
						Else
							ColumnNameEpiData = "({" & ColumnPrefix &  ColumnSectionNOffset & "0" & ColumnQuestionN-1 & "." & ColumnN & "})" & Space(2)
							SetExtendedAttribute oColumn, "NameEpiData", ColumnPrefix & ColumnSectionNOffset & "0" & ColumnQuestionN-1 & "." & ColumnN
						End If
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
						If Len(ColumnName)=0 Then
							Desc = Desc & ColumnNameEpiData & Space(NCharMax - oColumn.Length)
						ElseIf (Len(ColumnName)+ oColumn.Length) >= NCharMax Then
							Desc = Desc & Space(10) & ColumnName & ":" & vbCrLf
							Desc = Desc & Space(NCharWidth - NCharMax - 13)
							Desc = Desc & ColumnNameEpiData & Space(Max(NCharMax - oColumn.Length, 0))
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
		
		Desc = Replace(Desc, "\'", "'") 
		
		Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\" & LCase(oTable.Name) & ".qes", ForWriting, true)
		oFile.Write Desc & vbCrLf
		oFile.Close
		
	End If
Next

WScript.Echo "Création des fichiers CHK"

Desc = "LABELBLOCK" & vbCrLf

For Each oDomain in oModel.Domains
	If IsObject(oDomain) And (oDomain.LowValue="") And (oDomain.HighValue="") And (oDomain.ListOfValues <> "")Then
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

Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\header (labels).chk", ForWriting, true)
oFile.Write Desc & vbCrLf
oFile.Close

For Each oTable In oTables
   
   If IsObject(oTable) And (FormAll Or (oTable.Name=Form)) Then	
	
		WScript.Echo "  " & oTable.Name
				
		Desc =         "INCLUDE ""header.chk""" & vbCrLf
		Desc = Desc &  "INCLUDE ""header (labels).chk""" & vbCrLf & vbCrLf

		For Each oColumn in oTable.Columns
			SetExtendedAttribute oColumn, "Enabled", ""
		Next 
		
		For Each oColumn in oTable.Columns
			If IsObject(oColumn) And Not (oColumn.Computed) Then
				WScript.Echo "    Attribute: " & oColumn.Name & " (" & oColumn.Code & ")"

				ColumnName = UCase(Replace(ExtendedAttribute(oColumn, "NameEpiData"), ".", ""))
				
				S1 = ExtendedAttribute(oColumn, "Skip")
				S2 = ExtendedAttribute(oColumn, "Skip To")
				If (Len(S1)+Len(S2))>0 Then
					Flag = False		
					For Each oColumnInternal in oTable.Columns
					    If oColumnInternal.Code=S2 Then 
					    	Exit For
					    End If
					    If Flag Then
					    	SetExtendedAttribute oColumnInternal, "Enabled", ExtendedAttribute(oColumnInternal, "Enabled") & "(" & S1 & ") OR "
					    End if
				    	If oColumnInternal.Code=oColumn.Code Then Flag=True 			
					Next
				End If

			End If
		Next 

		For Each oColumn in oTable.Columns
			S = ExtendedAttribute(oColumn, "Enabled")
			If Len(S)>0 Then
				SetExtendedAttribute oColumn, "Enabled", "NOT (" & Mid(S, 1, Len(S)-4) & ")"
			End If
		Next 
		
		Desc = Desc & "BEFORE RECORD" & vbCrLf
		For Each oColumn in oTable.Columns
			If IsObject(oColumn) And Not (oColumn.Computed) Then
				S = ExtendedAttribute(oColumn, "Check")
				If InStr("NOENTER", S)>0 Then
					ColumnName = UCase(Replace(ExtendedAttribute(oColumn, "NameEpiData"), ".", ""))
					Desc = Desc & "  HIDE " & oColumn.Code & vbCrLf
				End If
			End If
		Next
		Desc = Desc & "END" & vbCrLf & vbCrLf

		'Desc = Desc & "AFTER RECORD" & vbCrLf
		'For Each oColumn in oTable.Columns
		'	If IsObject(oColumn) And Not (oColumn.Computed) Then
		'		
		'		WScript.Echo "    Attribute: " & oColumn.Name & " (" & oColumn.Code & ")"
'
'				ColumnName = UCase(Replace(ExtendedAttribute(oColumn, "NameEpiData"), ".", ""))
'				
'				Desc = Desc & "    IF (" & ColumnName & "=.) "
'				
'				S1 = ExtendedAttribute(oColumn, "Enabled")
'				
'				If Len(S1)>0 Then
'					Desc = Desc & "AND " & S1 & " THEN" & vbCrLf
'				Else
'				    Desc = Desc & "THEN" & vbCrLf 
'				End If
'				
'				Desc = Desc & "      HELP """ & ExtendedAttribute(oColumn, "NameEpiData") & " is mandatory.\n\nPlease check the data"" TYPE=WARNING" & vbCrLf
'				Desc = Desc & "      GOTO " & ColumnName & vbCrLf
'				Desc = Desc & "      EXIT" & vbCrLf
'				Desc = Desc & "    ENDIF" & vbCrLf 
'
'			End If
'		Next 
'		Desc = Desc & "END" & vbCrLf & vbCrLf
				
		For Each oColumn in oTable.Columns
			If IsObject(oColumn) And Not (oColumn.Computed) Then
				
				WScript.Echo "    Attribute: " & oColumn.Name & " (" & oColumn.Code & ")"
				Desc = Desc & vbCrLf
				Desc = Desc & "* " & ExtendedAttribute(oColumn, "NameEpiData") & " | "
				Desc = Desc & oColumn.Name & " (" & oColumn.Code & ")" & " | "
				Desc = Desc & oColumn.DataType & vbCrLf
				Desc = Desc & "* " & ExtendedAttribute(oColumn, "Question")
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
				If ExtendedAttribute(oColumn, "Skip") <> "" Then 
					Desc = Desc & "* Skip to """ & ExtendedAttribute(oColumn, "Skip To") & """ if (" & ExtendedAttribute(oColumn, "Skip") & ")" & vbCrLf
				Else
				    Desc = Desc & vbCrLf
				End If
				ColumnName = ColumnNameEpi(oColumn)
				If oColumn.Primary Then 
				    Desc = Desc & ExtendedAttribute(oTable, "Trigram") & vbCrLf		
				    Desc = Desc & "  KEY UNIQUE 1" & vbCrLf 
					Desc = Desc & "  NOENTER" & vbCrLf 
				ElseIf oColumn.ForeignKey Then
				    Desc = Desc & ColumnName & vbCrLf		
				    Desc = Desc & "  KEY 2" & vbCrLf
				    Desc = Desc & "  NOENTER" & vbCrLf 		
				ElseIf Not (oColumn.CannotModify) Then
					Desc = Desc & ColumnName & vbCrLf		
					If oColumn.Mandatory Then Desc = Desc & "  MUSTENTER" & vbCrLf
					If oColumn.LowValue<>"" And oColumn.HighValue<>"" Then 
						Desc = Desc & "  RANGE " & oColumn.LowValue & " " & oColumn.HighValue
						If oColumn.Domain.ListOfValues <> "" Then
							Values = oColumn.Domain.ListOfValues
							Values = Split(Values, vbNewLine, -1, 1)
							Desc = Desc & vbCrLf
							Desc = Desc & "  MISSINGVALUE"
							For i=0 To UBound(Values)
								If Values(i) <> "" Then
									Value = Values(i)
									Value = Split(Value, vbTab, -1, 1)
									Desc=Desc & " " & Value(0)
								End If
							Next	
						End If
						Desc = Desc & vbCrLf
						'Desc = Desc & "  BEFORE ENTRY" & vbCrLf
						'Desc = Desc & "    TYPE ""Allowed values between " & oColumn.LowValue & " and " & oColumn.HighValue
						'If oColumn.Domain.ListOfValues <> "" Then
					'	'	Desc = Desc & " (" & vbCrLf
						'	Values = oColumn.Domain.ListOfValues
						'	Values = Split(Values, vbNewLine, -1, 1)
						'	For i=0 To UBound(Values)
						'		If Values(i) <> "" Then
						'			Value = Values(i)
						'			Value = Split(Value, vbTab, -1, 1)
						'			Desc=Desc & Value(0) & ":" & Value(1) 
						'			If i < UBound(Values) Then Desc=Desc & ", "
						'		End If
						'	Next
						'	Desc = Desc & ")"
						'End If
						'Desc = Desc & """" & vbCrLf
						'Desc = Desc & "  END" & vbCrLf
					Else
						If oColumn.Domain.ListOfValues <> "" Then
							Desc = Desc & "  COMMENT LEGAL USE " & UCase(oColumn.Domain.Code) & vbCrLf
							Desc = Desc & "  TYPE COMMENT LEGAL" & vbCrLf
							
							'Values = oColumn.Domain.ListOfValues
							'Values = Split(Values, vbNewLine, -1, 1)
							'Desc = Desc & "  BEFORE ENTRY" & vbCrLf
							'Desc = Desc & "    TYPE ""Allowed values are "
							'For i=0 To UBound(Values)
						'		If Values(i) <> "" Then
						'			Value = Values(i)
						'			Value = Split(Value, vbTab, -1, 1)
						'			Desc=Desc & Value(0) 
						'			If i < UBound(Values) Then Desc=Desc & " or "
						'		End If
						'	Next
						'	Desc = Desc & """" & vbCrLf
						'	Desc = Desc & "  END" & vbCrLf
						End If
					End If
					S = ExtendedAttribute(oColumn, "Check")
					If Right(S, 3)="END" Then
					    S = Mid(S, 1, Len(S)-3)
				'		S = Replace(S, "      ", "  ¤¤¤¤")
						S = Replace(S, "     ", "  ¤¤¤")
						S = Replace(S, "    ", "  ¤¤")
						S = Replace(S, "   ", "  ¤")
						S = Replace(S, "  ", vbCrLf & "    ")
						S = Replace(S, "¤", " ")
						Desc = Desc & "  " & S & vbCrLf & "  END" & vbCrLf 
						For Each oBusinessRule in oColumn.AttachedRules
							If IsObject(oBusinessRule) Then
								'Desc = Desc & "  " & Replace(oBusinessRule.ServerExpression, "::", "  " & vbCrLf) & vbCrLf 
							End if
						Next
					ElseIf Right(S, 1)="¤" Then
					 	Desc = Desc & "  " & Replace(S, "¤", vbCrLf & "  ")
					ElseIf Len(S)>0 Then
					  Desc = Desc & "  " & S & vbCrLf 
					End If
					S1 = ExtendedAttribute(oColumn, "Skip")
					S2 = ExtendedAttribute(oColumn, "Skip To")
					If (Len(S1)+Len(S2))>0 Then
						Desc = Desc & "  AFTER ENTRY" & vbCrLf 
						Desc = Desc & "    IF (" & S1 & ") THEN" & vbCrLf 
						Desc = Desc & RepeatColumnCode(oTable, oColumn.Code, S2, "      HIDE %COLUMN%")
						Desc = Desc & RepeatColumnCode(oTable, oColumn.Code, S2, "      CLEAR %COLUMN%")
						Desc = Desc & "    ELSE" & vbCrLf 
						Desc = Desc & RepeatColumnCode(oTable, oColumn.Code, S2, "      UNHIDE %COLUMN%")				
						Desc = Desc & "    ENDIF" & vbCrLf 	
						Desc = Desc & "  END" & vbCrLf 	
					End If
				End If
				If oColumn = oTable.Columns.Item(oTable.Columns.Count - 1) Then
					Desc = Desc & "  GOTO WRITEREC" & vbCrLf	
				End if
				Desc = Desc & "END" & vbCrLf	
			End If
		Next 
		
		For Each oColumn in oTable.Columns
	       If IsObject(oColumn) Then
				ColumnName = UCase(Replace(ExtendedAttribute(oColumn, "NameEpiData"), ".", ""))
				Desc = Replace(Desc, oColumn.Code & "=", ColumnName & "=")
				Desc = Replace(Desc, oColumn.Code & "<", ColumnName & "<")
				Desc = Replace(Desc, oColumn.Code & ">", ColumnName & ">")
				Desc = Replace(Desc, oColumn.Code & " ", ColumnName & " ")
				Desc = Replace(Desc, oColumn.Code & ",", ColumnName & ",")
				Desc = Replace(Desc, oColumn.Code & "+", ColumnName & "+")
				Desc = Replace(Desc, oColumn.Code & """", ColumnName & """")
				Desc = Replace(Desc, oColumn.Code & vbCrLf, ColumnName & vbCrLf) 
			End If
		Next
		
		Desc = Replace(Desc, "\'", "'") 
		Desc = Replace(Desc, vbTab, "") 
		
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

Function Max(V1, V2)
  If V1<V2 Then
    Max = V2
  Else
    Max = V1
  End If
End Function

Function RepeatColumnCode(Table, ColumnCodeFrom, ColumnCodeTo, Source)
	Flag = False		
	RepeatColumnCode = ""
	For Each Column in Table.Columns
	    If Column.Code=ColumnCodeTo Then Exit For
	    If Flag Then
	    	RepeatColumnCode = RepeatColumnCode & Replace(Source, "%COLUMN%", ColumnNameEpi(Column)) & vbCrLf   
	    End if
    	If Column.Code=ColumnCodeFrom Then Flag=True 			
	Next
End Function

Function ColumnNameEpi(Column)
	ColumnNameEpi = UCase(Replace(ExtendedAttribute(Column, "NameEpiData"), ".", ""))
End Function




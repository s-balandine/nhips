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

WScript.Echo "Création des tables (fichiers)"

For Each oTable In oTables
   If IsObject(oTable) Then
      WScript.Echo "  " & oTable.Name
	  Desc = Desc & "/* SURVEY TITLE */" & vbCrLf
	  Desc = Desc & "/* FORM */" & vbCrLf
	  Desc = Desc & "/* .................................................................................... */" & vbCrLf & vbCrLf
      For Ni=0 to oTable.Columns.Count -1
         Set oColumn=oTable.Columns.Item(Ni)
         If IsObject(oColumn) And Not (oColumn.Computed) Then
            Desc = Desc & Space(4) & oColumn.Code &  Space(33-Len(oColumn.Code)) & oColumn.DataType
         End If
      Next
   
	  Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\" & LCase(oTable.Code) & ".qes", ForWriting, true)
	  oFile.Write Desc & vbCrLf
	  oFile.Close
	  
   End If
   
Next



Desc = vbCrLf
Desc = Desc & "/* Script DDL de création pour les vues */" & vbCrLf
Desc = Desc & "/* .................................................................................... */" & vbCrLf & vbCrLf

WScript.Echo "Création des vues"

For Each oView In oViews
   If IsObject(oView) Then
      Desc = Desc & oView.Preview & vbCrLf & vbCrLf
   End If
Next

Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\create views.sql", ForWriting, true)
oFile.Write Desc & vbCrLf
oFile.Close

Desc = vbCrLf
Desc = Desc & "/* Script DDL de création pour les contraintes de clé primaire */" & vbCrLf
Desc = Desc & "/* .................................................................................... */" & vbCrLf & vbCrLf

WScript.Echo "Création des contraintes de clé primaire"

For Each oTable In oTables
   For Each oKey In oTable.Keys
      If IsObject(oKey) And oKey.Primary Then
          WScript.Echo "  " & oKey.Name & " (" & oTable.Name & ")"
          Desc = Desc & oKey.Preview & vbCrLf
      End If
   Next
Next

Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\create keys primary.sql", ForWriting, true)
oFile.Write Desc & vbCrLf
oFile.Close

Desc = vbCrLf
Desc = Desc & "/* Script DDL de création pour les contraintes de clé étrangère */" & vbCrLf
Desc = Desc & "/* .................................................................................... */" & vbCrLf & vbCrLf

WScript.Echo "Création des contraintes de clé étrangère"

D = ""
For Each oTable In oTables
   For Each oReference In oTable.InReferences
      If IsObject(oReference)  Then
          WScript.Echo "  " & oReference.Name & " (" & oTable.Name & ")"
          D = D & oReference.Preview & vbCrLf
      End If
   Next
Next

Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\create keys foreign.sql", ForWriting, true)
oFile.Write Desc & Sort(D) & vbCrLf
oFile.Close

Desc = vbCrLf
Desc = Desc & "/* Script DDL de création pour les contraintes d'unicité */" & vbCrLf
Desc = Desc & "/* .................................................................................... */" & vbCrLf & vbCrLf

WScript.Echo "Création des contraintes d'unicité"

D = ""
For Each oTable In oTables
   For Each oIndex In oTable.Indexes
      If IsObject(oIndex) And Not (oIndex.Primary) And Not (oIndex.ForeignKey) And (oIndex.Unique)  Then
          WScript.Echo "  " & oIndex.Name & " (" & oTable.Name & ")"
          D = D & Replace(oIndex.Preview, vbCrLf, " ") & vbCrLf
      End If
   Next
Next

Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\create keys unique.sql", ForWriting, true)
oFile.Write Desc & Sort(D) & vbCrLf
oFile.Close

Desc = vbCrLf
Desc = Desc & "/* Script DDL de création pour les contraintes de champs */" & vbCrLf
Desc = Desc & "/* .................................................................................... */" & vbCrLf & vbCrLf

WScript.Echo "Création des contraintes de champs (default value)"

For Each oDomain In oDomains
   If IsObject(oDomain) And (oDomain.DefaultValue<>"") Then
      For Each oColumn In oDomain.Columns
        WScript.Echo "  " & oColumn.Name
        Desc = Desc & "ALTER TABLE " & oColumn.Table.Code & " ALTER COLUMN " & oColumn.Code & " SET DEFAULT '" & oDomain.DefaultValue & "'^" & vbCrLf
      Next
   End If
Next

Desc = Desc & vbCrLf

WScript.Echo "Création des contraintes de champs (min et max)"

For Each oDomain In oDomains
   If IsObject(oDomain) And (oDomain.LowValue<>"") Then
      For Each oColumn In oDomain.Columns
        WScript.Echo "  " & oColumn.Name
        Desc = Desc & "ALTER TABLE " & oColumn.Table.Code & " ALTER COLUMN " & oColumn.Code & " SET RANGE " & oDomain.LowValue & " " & oDomain.HighValue & "^" & vbCrLf
      Next
   End If
Next

Desc = Desc & vbCrLf

WScript.Echo "Création des contraintes de champs (code)"

For Each oDomain In oDomains
   If IsObject(oDomain) And (oDomain.ListOfValues<>"") Then
      Values = oDomain.ListOfValues
      Values = Split(Values, vbNewLine, -1, 1)
      For Each oColumn In oDomain.Columns
        WScript.Echo "  " & oColumn.Name
        Desc = Desc & "ALTER TABLE " & oColumn.Table.Code & " ALTER COLUMN " & oColumn.Code & " CHECK '"
        For i=0 To UBound(Values)
            Value = Values(i)
            Value = Split(Value, vbTab, -1, 1)
            Desc=Desc & Value(0)
            if i<UBound(Values) Then Desc=Desc & " OR "
        Next
        Desc = Desc & "'^" & vbCrLf
      Next
   End If
Next

Desc = Desc & vbCrLf

Set oFile = oFileSystemObject.OpenTextFile(strPathSql & "\create constraints fields.sql", ForWriting, true)
oFile.Write Desc
oFile.Close

Desc = vbCrLf
Desc = Desc & "/* Script DDL de création du schema de la base de données  */" & vbCrLf
Desc = Desc & "/* .................................................................................... */" & vbCrLf & vbCrLf

Desc = Desc & vbCrLf & "/* Informations générales */" & vbCrLf & vbCrLf

Desc = Desc & "CREATE TABLE T_S_DATABASE_INFO_DBI (" & vbCrLf
Desc = Desc & "  DBI_INFO  TEXT(128) NOT NULL," & vbCrLf
Desc = Desc & "  DBI_TYPE  TEXT(12)  NOT NULL," & vbCrLf
Desc = Desc & "  DBI_VALUE TEXT(255) NOT NULL," & vbCrLf
Desc = Desc & "  CONSTRAINT C_DBI_KP PRIMARY KEY (DBI_INFO))^" & vbCrLf & vbCrLf

Desc = Desc & "INSERT INTO T_S_DATABASE_INFO_DBI VALUES ('Language', 'TEXT', 'En')^" & vbCrLf
Desc = Desc & "INSERT INTO T_S_DATABASE_INFO_DBI VALUES ('Version' , 'TEXT', 'v1.0.0.0')^" & vbCrLf

Desc = Desc & vbCrLf & "/* Informations sur les tables */" & vbCrLf & vbCrLf

Desc = Desc & "CREATE TABLE T_S_TABLE_TBL ("
Desc = Desc & "  TBL_NAME TEXT(32)," & vbCrLf
Desc = Desc & "  CONSTRAINT C_TBL_KP PRIMARY KEY (TBL_NAME))^" & vbCrLf & vbCrLf

For Each oTable In oTables
  If IsObject(oTable) Then
    WScript.Echo "  " & oTable.Name
    Desc = Desc & "INSERT INTO T_S_TABLE_TBL VALUES (""" & oTable.Code & """)^" & vbCrLf
  End If
Next

Desc = Desc & vbCrLf & "/* Informations sur les champs */" & vbCrLf & vbCrLf

Desc = Desc & "CREATE TABLE T_S_FIELD_FLD (" & vbCrLf
Desc = Desc & "  FLD_TABLE          TEXT(32) NOT NULL," & vbCrLf
Desc = Desc & "  FLD_NAME           TEXT(32) NOT NULL," & vbCrLf
Desc = Desc & "  FLD_NAME_DBF       TEXT(32)," & vbCrLf
Desc = Desc & "  FLD_POSITION       SMALLINT NOT NULL," & vbCrLf
Desc = Desc & "  FLD_ATTRIBUTES     SMALLINT NOT NULL," & vbCrLf
Desc = Desc & "  FLD_DEFAULT_VALUE  TEXT(32)," & vbCrLf
Desc = Desc & "  FLD_REQUIRED       LOGICAL NOT NULL," & vbCrLf
Desc = Desc & "  FLD_READONLY       LOGICAL NOT NULL," & vbCrLf
Desc = Desc & "  FLD_SIZE           SMALLINT NOT NULL," & vbCrLf
Desc = Desc & "  FLD_TYPE           TEXT(32) NOT NULL," & vbCrLf
Desc = Desc & "  FLD_VALIDATION_RULE TEXT(100)," & vbCrLf
Desc = Desc & "  FLD_VALIDATION_TEXT TEXT(255)," & vbCrLf
Desc = Desc & "  CONSTRAINT C_FLD_KP PRIMARY KEY (FLD_TABLE, FLD_NAME))^" & vbCrLf & vbCrLf

Desc = Desc & "CREATE TABLE T_S_FIELD_TRANSLATION_FTL (" & vbCrLf
Desc = Desc & "  FTL_TABLE          TEXT(32)," & vbCrLf
Desc = Desc & "  FTL_NAME           TEXT(32) NOT NULL," & vbCrLf
Desc = Desc & "  FTL_NAME_UI        TEXT(32) NOT NULL," & vbCrLf
Desc = Desc & "  FTL_LANGUAGE       TEXT(2) NOT NULL," & vbCrLf
Desc = Desc & "  CONSTRAINT C_FTL_KP PRIMARY KEY (FTL_TABLE, FTL_NAME, FTL_LANGUAGE))^" & vbCrLf & vbCrLf

Desc = Desc & "CREATE VIEW V_S_FIELD_FLD AS" & vbCrLf
Desc = Desc & "  PARAMETERS [LANG] TEXT(2);" & vbCrLf
Desc = Desc & "  SELECT" & vbCrLf
Desc = Desc & "    FLD_TABLE," & vbCrLf
Desc = Desc & "    FLD_NAME," & vbCrLf
Desc = Desc & "    FTL_NAME_UI AS FLD_NAME_UI," & vbCrLf
Desc = Desc & "    FLD_NAME_DBF," & vbCrLf
Desc = Desc & "    FLD_POSITION," & vbCrLf
Desc = Desc & "    FLD_ATTRIBUTES," & vbCrLf
Desc = Desc & "    FLD_DEFAULT_VALUE," & vbCrLf
Desc = Desc & "    FLD_REQUIRED," & vbCrLf
Desc = Desc & "    FLD_READONLY," & vbCrLf
Desc = Desc & "    FLD_SIZE," & vbCrLf
Desc = Desc & "    FLD_TYPE," & vbCrLf
Desc = Desc & "    FLD_VALIDATION_RULE," & vbCrLf
Desc = Desc & "    FLD_VALIDATION_TEXT" & vbCrLf
Desc = Desc & "  FROM" & vbCrLf
Desc = Desc & "    T_S_FIELD_FLD," & vbCrLf
Desc = Desc & "    T_S_FIELD_TRANSLATION_FTL" & vbCrLf
Desc = Desc & "  WHERE" & vbCrLf
Desc = Desc & "    (FLD_TABLE = FTL_TABLE) AND" & vbCrLf
Desc = Desc & "    (FLD_NAME  = FTL_NAME) AND" & vbCrLf
Desc = Desc & "    (FTL_LANGUAGE=[LANG])^" & vbCrLf & vbCrLf
    
For Each oTable In oTables
  If IsObject(oTable) Then

    WScript.Echo "  " & oTable.Name

    Index=0

    For Each oColumn In oTable.Columns
      If IsObject(oColumn) Then

        WScript.Echo "    " & oColumn.Name

        Desc = Desc & "INSERT INTO T_S_FIELD_FLD             VALUES ("

        Desc = Desc &   """" & oColumn.Table.Code & """, "
        Desc = Desc &   """" & oColumn.Code & """, "
        Desc = Desc &   """" & ExtendedAttribute(oColumn, "ExtFieldDbaseLabel") & """, "
        Desc = Desc &   Index & ", "
        Desc = Desc &   "0, "
        Desc = Desc &   """" & oColumn.DefaultValue & """, "
        If oColumn.Mandatory Then
          Desc = Desc & "TRUE, "
        Else
          Desc = Desc & "FALSE, "
        End If
        If oColumn.CannotModify Then
          Desc = Desc & "TRUE, "
        Else
          Desc = Desc & "FALSE, "
        End If
        Select Case oColumn.Datatype
        Case "BOOLEAN":
          Desc = Desc &  "1, "
        Case "SMALLINT":
          Desc = Desc &  "2, "
        Case "LONG":
          Desc = Desc &  "4, "
        Case "AUTOINCREMENT":
          Desc = Desc &  "4, "
        Case "FLOAT":
          Desc = Desc &  "4, "
        Case "SINGLE":
          Desc = Desc &  "4, "
        Case "DATETIME":
          Desc = Desc &  "8, "
        Case "TEXT":
          Desc = Desc &   oColumn.Length & ", "
        Case Else
          Desc = Desc &   oColumn.Length & ", "
        End Select

        If oColumn.Primary Then
          Desc = Desc &  """AUTOINCREMENT"", "
        Else
          Select Case oColumn.Datatype
          Case "TEXT":
            Desc = Desc &   """" & oColumn.Datatype & "(" & oColumn.Length & ")"", "
          Case Else
            Desc = Desc &   """" & oColumn.Datatype & """, "
          End Select
        End If

        ValidationRule = ""

        If Not oColumn.Computed and oColumn.name<>"Key Section" and oColumn.name<>"Key Name" and oColumn.name<>"Key Value" Then
          If not oColumn.Domain is nothing Then
            If oColumn.Domain.ListOfValues <> "" Then
              Values = oColumn.Domain.ListOfValues
              Values = Split(Values, vbNewLine, -1, 1)
              For i=0 To UBound(Values)
                Value = Values(i)
                Value = Split(Value, vbTab, -1, 1)
                ValidationRule=ValidationRule & Value(0)
                if i<UBound(Values) Then ValidationRule=ValidationRule & " OR "
              Next
            End If
          End If
        End If

        Desc = Desc &  """" & ValidationRule & """, "

        Desc = Desc &  """" & ExtendedAttribute(oColumn, "ExtFieldDisplayLabel")
        Select Case oColumn.Datatype
        Case "BOOLEAN":
          Desc = Desc &  " is a logical field"
        Case "SMALLINT":
          Desc = Desc &  " is a small integer field"
        Case "LONG":
          Desc = Desc &  " is a integer field"
        Case "AUTOINCREMENT":
          Desc = Desc &  " is a autoincrement field"
        Case "FLOAT":
          Desc = Desc &  " is a float field"
        Case "SINGLE":
          Desc = Desc &  " is a single field"
        Case "DATETIME":
          Desc = Desc &  " is a datetime field"
        Case "TEXT":
          Desc = Desc &  " is a text field of " & oColumn.Length & " char(s)"
        Case Else
          Desc = Desc &  " is a text field of " & oColumn.Length & " char(s)"
        End Select
        If oColumn.Mandatory Then Desc = Desc &  " mandatory"
        If oColumn.CannotModify Then Desc = Desc &  " read-only"
        Desc = Desc & """)^" & vbCrLf

        Languages=Array("En", "Fr")
        For Each Language In Languages
          Desc = Desc & "INSERT INTO T_S_FIELD_TRANSLATION_FTL VALUES ("
          Desc = Desc &  """" & oColumn.Table.Code & """, "
          Desc = Desc &  """" & oColumn.Code & """, "
          Desc = Desc &  """" & ExtendedAttribute(oColumn, "ExtFieldDisplayLabel") & """, "
          Desc = Desc &  """" & Language & """)^" & vbCrLf
        Next
        
        Index=Index+1

      End If
    Next

  End If
Next

Desc = Desc & vbCrLf & "/* Informations sur les codes */" & vbCrLf & vbCrLf

Desc = Desc & "CREATE TABLE T_S_CODE_CDE (" & vbCrLf
Desc = Desc & "  CDE_GROUP             TEXT(32)," & vbCrLf
Desc = Desc & "  CDE_CODE              SMALLINT NOT NULL," & vbCrLf
Desc = Desc & "  CDE_DESCRIPTION_SHORT TEXT(32) NOT NULL," & vbCrLf
Desc = Desc & "  CONSTRAINT C_CDE_KP PRIMARY KEY (CDE_GROUP, CDE_CODE))^" & vbCrLf & vbCrLf

Desc = Desc & "CREATE TABLE T_S_CODE_TRANSLATION_CTL (" & vbCrLf
Desc = Desc & "  CTL_GROUP       TEXT(32)," & vbCrLf
Desc = Desc & "  CTL_CODE        SMALLINT NOT NULL," & vbCrLf
Desc = Desc & "  CTL_DESCRIPTION TEXT(32) NOT NULL," & vbCrLf
Desc = Desc & "  CTL_LANGUAGE    TEXT(2) NOT NULL," & vbCrLf
Desc = Desc & "  CONSTRAINT C_CDE_KP PRIMARY KEY (CTL_GROUP, CTL_CODE, CTL_LANGUAGE))^" & vbCrLf & vbCrLf

Desc = Desc & "CREATE VIEW V_S_CODE_CDE AS" & vbCrLf
Desc = Desc & "  PARAMETERS [LANG] TEXT(2);" & vbCrLf
Desc = Desc & "  SELECT" & vbCrLf
Desc = Desc & "    CDE_GROUP," & vbCrLf
Desc = Desc & "    CDE_CODE," & vbCrLf
Desc = Desc & "    CDE_DESCRIPTION_SHORT," & vbCrLf
Desc = Desc & "    CTL_DESCRIPTION AS CDE_DESCRIPTION" & vbCrLf
Desc = Desc & "  FROM" & vbCrLf
Desc = Desc & "    T_S_CODE_CDE," & vbCrLf
Desc = Desc & "    T_S_CODE_TRANSLATION_CTL" & vbCrLf
Desc = Desc & "  WHERE" & vbCrLf
Desc = Desc & "    (CDE_GROUP = CTL_GROUP) AND" & vbCrLf
Desc = Desc & "    (CDE_CODE  = CTL_CODE) AND" & vbCrLf
Desc = Desc & "    (FTL_LANGUAGE=[LANG])^" & vbCrLf & vbCrLf

WScript.Echo vbCrLf & "Domains (Code)"

For Each oDomain In oDomains
   If IsObject(oDomain) Then
    If oDomain.ListOfValues <> "" Then
    
     WScript.Echo "  " & oDomain.Name
     
     d = oDomain.ListOfValues
     d = Replace(d, vbtab, "¤")
     d = Replace(d, vbNewLine, "^")
     d = d & "^"
     d1 = ""
     d2 = ""
     d3 = ""
     d4 = ""
     d5 = ""
     do while (InStr(d, "¤") > 0)
      p1 = InStr(d, "¤")
      p2 = InStr(d, "|")
      p3 = InStr(d, "^")
      d0 = Mid(d, p1 + 1, p3 - p1 - 1)
      d0 = d0 & "|||||"
      d1 = Mid(d0, 1, InStr(d0, "|") - 1)
      d0 = Mid(d0, InStr(d0, "|") + 1)
      d2 = Mid(d0, 1, InStr(d0, "|") - 1)
      d0 = Mid(d0, InStr(d0, "|") + 1)
      d3 = Mid(d0, 1, InStr(d0, "|") - 1)
      d0 = Mid(d0, InStr(d0, "|") + 1)
      d4 = Mid(d0, 1, InStr(d0, "|") - 1)
      d0 = Mid(d0, InStr(d0, "|") + 1)
      d5 = Mid(d0, 1, InStr(d0, "|") - 1)
      
      Desc = Desc & "INSERT INTO T_S_CODE_CDE             VALUES ('" & oDomain.Code & "','" & Mid(d, 1, p1 - 1) & "','" & d1 & "')^" & vbCrLf
      If d2<>"" Then Desc = Desc & "INSERT INTO T_S_CODE_TRANSLATION_CTL VALUES ('" & oDomain.Code & "','" & Mid(d, 1, p1 - 1) & "','" & Replace(d2, "'", "''") & "', 'En')^" & vbCrLf
      If d3<>"" Then Desc = Desc & "INSERT INTO T_S_CODE_TRANSLATION_CTL VALUES ('" & oDomain.Code & "','" & Mid(d, 1, p1 - 1) & "','" & Replace(d3, "'", "''") & "', 'Fr')^" & vbCrLf
   
      d = Mid(d, p3 + 1)
     loop
    end if
   End If
Next

Set File = oFileSystemObject.OpenTextFile(strPathSQL & "\create schema.sql", ForWriting, true)
File.Write Desc & vbCrLf
File.Close

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
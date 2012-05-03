'******************************************************************************
'* File:     generate.vbs
'* Purpose:  Generate wiki reference from model
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

StrModelId = "epitryps"
StrModel = ""
StrSpace = ""

If oArgs.count()>0 Then
  StrModel = oArgs(0)
  StrSpace = oArgs(1)
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

If StrModel="" Then WScript.Quit

If Not IsObject(ActiveModel) Then
  Set oApp = CreateObject("PowerDesigner.Application")
  Set oModel = oApp.OpenModel(StrModel,0)
Else
  Set oModel = ActiveModel
  strPathModel = oFileSystemObject.GetParentFolderName(ActiveModel.Filename)
End If

strPathModel = CreateFolder(oFileSystemObject.GetParentFolderName(oModel.Filename))
strPathWiki = CreateFolder(strPathModel & "\help")
strPathWiki = CreateFolder(strPathWiki & "\wiki")

WScript.Echo "Begin script generation"

Wiki = ""

'ActiveDiagram.ExportImage(strPathWiki & "\diagram.png")

SortCollection(oModel.Domains)
SortCollection(oModel.Tables)

' Liste des domaines
' .......................................

WScript.Echo "Liste des domaines..."

Desc = Style
Desc = Desc &   "{table:class=register}" & vbCrLf
Desc = Desc &   "  {thead}" & vbCrLf
Desc = Desc &   "    {tr:class=h1}" & vbCrLf
Desc = Desc &   "      {td:colspan=2}Domaine{td}" & vbCrLf
Desc = Desc &   "      {td}Type{td}" & vbCrLf
Desc = Desc &   "      {td}M{td}" & vbCrLf
Desc = Desc &   "      {td}Defaut{td}" & vbCrLf
Desc = Desc &   "      {td}Description{td}" & vbCrLf
Desc = Desc &   "    {tr}" & vbCrLf
Desc = Desc &   "  {thead}" & vbCrLf
Desc = Desc &   "  {tbody}" & vbCrLf

For Each oDomain In oModel.Domains
  If IsObject(oDomain) Then
    WScript.Echo oDomain

    Desc = Desc &   "    {tr}" & vbCrLf
    Desc = Desc &   "      {td}[" &       oDomain.Name & "|" & oDomain.Name & " (Domain)]{td}" & vbCrLf
    Desc = Desc &   "      {td}"  &       oDomain.Code & "{td}" & vbCrLf
    Desc = Desc &   "      {td}"  &       oDomain.Datatype & "{td}" & vbCrLf
    Desc = Desc &   "      {td}"  & YesNo(oDomain.Mandatory) & "{td}" & vbCrLf
    Desc = Desc &   "      {td}"  &       oDomain.DefaultValue & "{td}" & vbCrLf
    Desc = Desc &   "      {td}"  &       oDomain.Comment & "{td}" & vbCrLf
    Desc = Desc &   "    {tr}" & vbCrLf

  End If
Next

Desc = Desc &   "  {tbody}" & vbCrLf
Desc = Desc & "{table}" & vbCrLf

Set File = oFileSystemObject.OpenTextFile(strPathWiki & "\domains.wiki", ForWriting, true)
File.Write Desc & vbCrLf
File.Close

Wiki = Wiki + "domains ""Liste des domaines"" ""Modele de donnees"" " & StrSpace & vbCrLf

' Pour chaque domaine
' .......................................

strPathWikiDomain = CreateFolder(strPathWiki & "\domains")

WScript.Echo "Description des domaines..."

For Each oDomain In oModel.Domains
  If IsObject(oDomain) Then
    WScript.Echo oDomain
    
    Desc = Style
    Desc = Desc &   "{set-data:name|hidden=true}" & oDomain.Name & "{set-data}" & vbCrLf
    Desc = Desc &   "{set-data:name-link|hidden=true}" & oDomain.Name & " (Domain){set-data}" & vbCrLf
    Desc = Desc &   "{set-data:code|hidden=true}" & oDomain.Code & "{set-data}" & vbCrLf
    Desc = Desc &   "{set-data:datatype|hidden=true}" & oDomain.Datatype & "{set-data}" & vbCrLf
    Desc = Desc &   "{set-data:is-required|hidden=true}" & YesNo(oDomain.Mandatory) & "{set-data}" & vbCrLf
    Desc = Desc &   "{set-data:default|hidden=true}" & oDomain.DefaultValue & "{set-data}" & vbCrLf
    Desc = Desc &   "{set-data:comment|hidden=true}" & oDomain.Comment & "{set-data}" & vbCrLf

    Desc = Desc &   "{section}" & vbCrLf

    Desc = Desc &   "  {column:width=20%}" & vbCrLf

    Desc = Desc &   "    {panel}{navigator}{panel}" & vbCrLf

    Desc = Desc &   "    {panel:title=Sommaire}" & vbCrLf
    Desc = Desc &   "      {toc:indent=10px|maxLevel=3}" & vbCrLf
    Desc = Desc &   "    {panel}" & vbCrLf

    Desc = Desc &   "    {info:title=Information technique|icon=false}" & vbCrLf
    Desc = Desc &   "      ----" & vbCrLf
    Desc = Desc &   "      {table:class=spec}" & vbCrLf
    Desc = Desc &   "        {tr}" & vbCrLf
    Desc = Desc &   "          {td}*Nom*:{td}" & vbCrLf
    Desc = Desc &   "          {td}{get-data:name}{td}" & vbCrLf
    Desc = Desc &   "        {tr}" & vbCrLf
    Desc = Desc &   "        {tr}" & vbCrLf
    Desc = Desc &   "          {td}&nbsp;&nbsp;SQL:{td}" & vbCrLf
    Desc = Desc &   "          {td}{get-data:code}{td}" & vbCrLf
    Desc = Desc &   "        {tr}" & vbCrLf
    Desc = Desc &   "        {tr}" & vbCrLf
    Desc = Desc &   "          {td}*Type*:{td}" & vbCrLf
    Desc = Desc &   "          {td}{get-data:datatype}{td}" & vbCrLf
    Desc = Desc &   "        {tr}" & vbCrLf
    Desc = Desc &   "        {tr}" & vbCrLf
    Desc = Desc &   "          {td}*Obligatoire*:{td}" & vbCrLf
    Desc = Desc &   "          {td:style=float:left;}{get-data:is-required}{td}" & vbCrLf
    Desc = Desc &   "        {tr}" & vbCrLf
    Desc = Desc &   "      {table}" & vbCrLf
    Desc = Desc &   "    {info}" & vbCrLf

    Desc = Desc &   "  {column}" & vbCrLf

    Desc = Desc &   "  {column}" & vbCrLf
    
    Desc = Desc &   "h2. Définition" & vbCrLf & vbCrLf
    Desc = Desc &   oDomain.Comment & vbCrLf & vbCrLf
    Desc = Desc &   "h2. Raison et cas d'utilisation" & vbCrLf & vbCrLf
    Desc = Desc &   "..." & vbCrLf  & vbCrLf
    Desc = Desc &   "h2. Discussion" & vbCrLf & vbCrLf
    Desc = Desc &   "..." & vbCrLf  & vbCrLf
    Desc = Desc &   "h2. Contraintes " & vbCrLf & vbCrLf

    If Mid(oDomain.datatype,1,7) = "NUMERIC" Then
      If oDomain.LowValue <> oDomain.HighValue Then
        If Note<>"" Then Note = Note & " \\ \\ "
        Note = Note &  "Seules les valeurs entre *" & oDomain.LowValue & "* et *"  & oDomain.HighValue & "* sont permises."
      End if
    End If

    Note = ""
    If oDomain.name<>"Key Section" and oDomain.name<>"Key Name" and oDomain.name<>"Key Value" Then
      If oDomain.ListOfValues <> "" Then
        If Note<>"" Then Note = Note & " \\ \\ "
        Note = CodeTab(oDomain.ListOfValues, oDomain.DefaultValue)
      End If
    End If

    Desc = Desc & Note & vbCrLf
    
    Desc = Desc &   "    h2. Liste des colonnes du domaine" & vbCrLf
    Desc = Desc &   "    {table:class=register|width=50%}" & vbCrLf
    Desc = Desc &   "      {thead}" & vbCrLf
    Desc = Desc &   "        {tr:class=h1}" & vbCrLf
    Desc = Desc &   "          {td}Table{td}" & vbCrLf
    Desc = Desc &   "          {td}Variable{td}" & vbCrLf
    Desc = Desc &   "          {td}Variable (SQL){td}" & vbCrLf
    Desc = Desc &   "        {tr}" & vbCrLf
    Desc = Desc &   "      {thead}" & vbCrLf
    Desc = Desc &   "      {tbody}" & vbCrLf

    For Each oColumn In oDomain.Columns
      If IsObject(oColumn) Then
        WScript.Echo oColumn

        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}[" &       oColumn.Table.Name & "]{td}" & vbCrLf
        Desc = Desc &   "          {td}[" &       oColumn.Name & "|" & oColumn.Name & " (" & oColumn.Table.Stereotype & ")]{td}" & vbCrLf
        Desc = Desc &   "          {td}"  &       oColumn.Code & "{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf

      End If
    Next

    Desc = Desc &   "      {tbody}" & vbCrLf
    Desc = Desc &   "    {table}" & vbCrLf

    Desc = Desc &   "  {column}" & vbCrLf

    Desc = Desc &   "{section}" & vbCrLf

    Set File = oFileSystemObject.OpenTextFile(strPathWikiDomain & "\" & lcase(oDomain.Code) & ".wiki", ForWriting, true)
    File.Write Desc & vbCrLf
    File.Close

    Wiki = Wiki & "domains" & "/" & lcase(oDomain.Code) & " """ & oDomain.Name & " (Domain)"" ""Liste des domaines"" " & StrSpace & vbCrLf

  End If
Next

' Liste des tables
' .......................................

WScript.Echo "Liste des tables..."

Desc = Style
Desc = Desc &   "{table:class=register}" & vbCrLf
Desc = Desc &   "  {thead}" & vbCrLf
Desc = Desc &   "    {tr:class=h1}" & vbCrLf
Desc = Desc &   "      {td:colspan=2}Table{td}" & vbCrLf
Desc = Desc &   "      {td}Description{td}" & vbCrLf
Desc = Desc &   "    {tr}" & vbCrLf
Desc = Desc &   "  {thead}" & vbCrLf
Desc = Desc &   "  {tbody}" & vbCrLf

For Each oTable In oModel.Tables
  If IsObject(oTable) Then
    WScript.Echo oTable.Code

    Desc = Desc &   "    {tr}" & vbCrLf
    Desc = Desc &   "      {td}[" &     oTable.Name & "]{td}" & vbCrLf
    Desc = Desc &   "      {td}"  &     oTable.Code & "{td}" & vbCrLf
    Desc = Desc &   "      {td}"  &     oTable.Comment & "{td}" & vbCrLf
    Desc = Desc &   "    {tr}" & vbCrLf

  End If
Next

Desc = Desc &   "  {tbody}" & vbCrLf
Desc = Desc & "{table}" & vbCrLf

Set File = oFileSystemObject.OpenTextFile(strPathWiki & "\tables.wiki", ForWriting, true)
File.Write Desc & vbCrLf
File.Close

Wiki = Wiki + "tables ""Liste des tables"" ""Modele de donnees"" " & StrSpace & vbCrLf

' Pour chaque tables
' .......................................

strPathWikiTable = CreateFolder(strPathWiki & "\tables")
 
WScript.Echo "Tables..."

For Each oTable In oModel.Tables
  If IsObject(oTable) Then
    WScript.Echo oTable.Code

    Desc = Style
    Desc = Desc &   "h2. Définition" & vbCrLf & vbCrLf
    Desc = Desc &   oTable.Comment & vbCrLf
    Desc = Desc &   "h2. Champs" & vbCrLf & vbCrLf
    Desc = Desc &   "{table:class=register}" & vbCrLf
    Desc = Desc &   "  {thead}" & vbCrLf
    Desc = Desc &   "    {tr:class=h1}" & vbCrLf
    Desc = Desc &   "      {td:colspan=2}Variable{td}" & vbCrLf
    Desc = Desc &   "      {td}Domaine{td}" & vbCrLf
    Desc = Desc &   "      {td}Type{td}" & vbCrLf
    Desc = Desc &   "      {td}PK{td}" & vbCrLf
    Desc = Desc &   "      {td}FK{td}" & vbCrLf
    Desc = Desc &   "      {td}NN{td}" & vbCrLf
    Desc = Desc &   "      {td}Description{td}" & vbCrLf
    Desc = Desc &   "    {tr}" & vbCrLf
    Desc = Desc &   "  {thead}" & vbCrLf
    Desc = Desc &   "  {tbody}" & vbCrLf

    For Each oColumn In oTable.Columns
      If IsObject(oColumn) Then
        WScript.Echo oColumn.Code

        Desc = Desc &   "    {tr}" & vbCrLf
        Desc = Desc &   "      {td}[" & oColumn.Name & "|" & oColumn.Name & " (" & lcase(oTable.Stereotype) & ")" & "]{td}" & vbCrLf
        Desc = Desc &   "      {td}"  & oColumn.Code & "{td}" & vbCrLf

        If Not oColumn.Computed and oColumn.name<>"Key Section" and oColumn.name<>"Key Name" and oColumn.name<>"Key Value" Then
          If oColumn.Domain is nothing Then
            Desc = Desc &   "      {td}[]{td}" & vbCrLf
          Else
            Desc = Desc &   "      {td}["  &    oColumn.Domain.Name & "|" & oColumn.Domain.Name & " (Domain)]{td}" & vbCrLf
          End If
        Else
          Desc = Desc &   "      {td}Calculé{td}" & vbCrLf
        End If

        Desc = Desc &   "      {td}"  &         oColumn.DataType & "{td}" & vbCrLf
        Desc = Desc &   "      {td}"  & YesNo(  oColumn.Primary) & "{td}" & vbCrLf
        Desc = Desc &   "      {td}"  & YesNo(  oColumn.ForeignKey) & "{td}" & vbCrLf
        Desc = Desc &   "      {td}"  & YesNo(  oColumn.Mandatory) & "{td}" & vbCrLf

        Note = oColumn.Comment

        if oColumn.ForeignKey then
           if oColumn.comment <>"" then desc1 = desc1 & "<br><br>"
           for each oForeignKeyJoin in oColumn.ForeignKeyJoins
             If Note<>"" Then Note = Note & " \\ \\ "
             Note = Note & "Clé vers la table [" & oForeignKeyJoin.ParentTableColumn.Table.Name & "]"
           next
        end if

        If Mid(oColumn.datatype,1,7) = "NUMERIC" Then
          If oColumn.LowValue <> oColumn.HighValue Then
            If Note<>"" Then Note = Note & " \\ \\ "
              Note = Note &  "Seules les valeurs entre *" & oColumn.LowValue & "* et *"  & oColumn.HighValue & "* sont permises."
            End if
          End If

        If Not oColumn.Computed and oColumn.name<>"Key Section" and oColumn.name<>"Key Name" and oColumn.name<>"Key Value" Then
          If not oColumn.Domain is nothing Then
            If oColumn.Domain.ListOfValues <> "" Then
              If Note<>"" Then Note = Note & " \\ \\ "
              Note = CodeTabSimple(oColumn.Domain.ListOfValues, oColumn.DefaultValue)
            End If
          End If
        End If

        Desc = Desc &   "      {td}"  &     Note & "{td}" & vbCrLf
        Desc = Desc &   "    {tr}" & vbCrLf

      End If
    Next

    Desc = Desc &   "  {tbody}" & vbCrLf
    Desc = Desc & "{table}" & vbCrLf

    Set File = oFileSystemObject.OpenTextFile(strPathWikiTable & "\" & lcase(oTable.Code) & ".wiki", ForWriting, true)
    File.Write Desc & vbCrLf
    File.Close

    Wiki = Wiki & "tables" & "/" & lcase(oTable.Code) & " """ & oTable.Name & """ ""Liste des tables"" " & StrSpace & vbCrLf
  
  End If
Next

' Pour chaque variables
' .......................................

WScript.Echo "Liste des variables..."

Desc = Style
Desc = Desc &   "{table-plus:sortIcon=true}" & vbCrLf & vbCrLf
Desc = Desc &   " || Table || Variable || Nom SQL || Nom || Nom Export (Stata) ||" & vbCrLf

For Each oTable In oModel.Tables
  If IsObject(oTable) Then

    For Each oColumn In oTable.Columns
      If IsObject(oColumn) Then

        Desc = Desc &   " | [" & oTable.Name  & "]"
        Desc = Desc &   " | [" & oColumn.Name & "]"
        Desc = Desc &   " |  " & oColumn.Code
        Desc = Desc &   " |  " & oColumn.getextendedattribute(oModel.DBMS.Code & ".ExtFieldDisplayLabel")
        Desc = Desc &   " |  " & oColumn.getextendedattribute(oModel.DBMS.Code & ".ExtFieldDbaseLabel") & " |" & vbCrLf

      End If
    Next

  End If
Next

Desc = Desc &   "{table-plus}" & vbCrLf & vbCrLf

Set File = oFileSystemObject.OpenTextFile(strPathWiki & "\fields.wiki", ForWriting, true)
File.Write Desc & vbCrLf
File.Close

Wiki = Wiki + "fields ""Liste des variables"" ""Modele de donnees"" " & StrSpace & vbCrLf

WScript.Echo "Descriptifs des variables..."

For Each oTable In oModel.Tables
  If IsObject(oTable) Then
    WScript.Echo oTable.Code

    CreateFolder(strPathWiki & "\tables\" & lcase(oTable.Code))

    For Each oColumn In oTable.Columns
      If IsObject(oColumn) Then

        Desc = Style
        Desc = Desc &   "{set-data:table|hidden=true}" & oColumn.Table.Name & "{set-data}" & vbCrLf
        Desc = Desc &   "{set-data:name|hidden=true}" & oColumn.Name & "{set-data}" & vbCrLf
        Desc = Desc &   "{set-data:name-link|hidden=true}" & oColumn.Name & " (Domain){set-data}" & vbCrLf
        Desc = Desc &   "{set-data:code|hidden=true}" & oColumn.Code & "{set-data}" & vbCrLf
        Desc = Desc &   "{set-data:code-stata|hidden=true}" & Alias(oColumn.getextendedattribute(oModel.DBMS.Code & ".ExtFieldDbaseLabel")) & "{set-data}" & vbCrLf
        Desc = Desc &   "{set-data:code-common|hidden=true}" & Alias(oColumn.getextendedattribute(oModel.DBMS.Code & ".ExtFieldDisplayLabel")) & "{set-data}" & vbCrLf
        Desc = Desc &   "{set-data:datatype|hidden=true}" & oColumn.Datatype & "{set-data}" & vbCrLf
        Desc = Desc &   "{set-data:is-primary|hidden=true}" & YesNo(oColumn.Primary) & "{set-data}" & vbCrLf
        Desc = Desc &   "{set-data:is-foreign|hidden=true}" & YesNo(oColumn.ForeignKey) & "{set-data}" & vbCrLf
        Desc = Desc &   "{set-data:is-required|hidden=true}" & YesNo(oColumn.Mandatory) & "{set-data}" & vbCrLf
        Desc = Desc &   "{set-data:domain|hidden=true}" & oColumn.Domain.Name & "{set-data}" & vbCrLf

        Desc = Desc &   "{section}" & vbCrLf

        Desc = Desc &   "  {column:width=20%}" & vbCrLf

        Desc = Desc &   "    {panel}{navigator}{panel}" & vbCrLf

        Desc = Desc &   "    {panel:title=Sommaire}" & vbCrLf
        Desc = Desc &   "      {toc:indent=10px|maxLevel=3}" & vbCrLf
        Desc = Desc &   "    {panel}" & vbCrLf

        Desc = Desc &   "    {info:title=Information technique|icon=false}" & vbCrLf
        Desc = Desc &   "      ----" & vbCrLf
        Desc = Desc &   "      {table:class=spec}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}*Table*:{td}" & vbCrLf
        Desc = Desc &   "          {td}[" & oColumn.Table.Name & "]{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}*Nom*:{td}" & vbCrLf
        Desc = Desc &   "          {td}{get-data:name}{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}&nbsp;&nbsp;Interface:{td}" & vbCrLf
        Desc = Desc &   "          {td}{get-data:code-common}{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}&nbsp;&nbsp;SQL:{td}" & vbCrLf
        Desc = Desc &   "          {td}{get-data:code}{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}&nbsp;&nbsp;Stata:{td}" & vbCrLf
        Desc = Desc &   "          {td}{get-data:code-stata}{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}*Domain*:{td}" & vbCrLf
        Desc = Desc &   "          {td}[" & oColumn.Domain.Name & "|" & oColumn.Domain.Name & " (Domain)]{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}&nbsp;&nbsp;Type:{td}" & vbCrLf
        Desc = Desc &   "          {td}{get-data:datatype}{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}*Clé*:{td}" & vbCrLf
        Desc = Desc &   "          {td}{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}&nbsp;&nbsp;Clé&nbsp;primaire:{td}" & vbCrLf
        Desc = Desc &   "          {td:style=float:left;}{get-data:is-primary}{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}&nbsp;&nbsp;Clé&nbsp;étrangère:{td}" & vbCrLf
        Desc = Desc &   "          {td:style=float:left;}{get-data:is-foreign}{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "          {td}*Obligatoire*:{td}" & vbCrLf
        Desc = Desc &   "          {td:style=float:left;}{get-data:is-required}{td}" & vbCrLf
        Desc = Desc &   "        {tr}" & vbCrLf
        Desc = Desc &   "      {table}" & vbCrLf
        Desc = Desc &   "    {info}" & vbCrLf

        Desc = Desc &   "  {column}" & vbCrLf

        Desc = Desc &   "  {column}" & vbCrLf
        
        Desc = Desc &   "h2. Définition" & vbCrLf & vbCrLf
        Desc = Desc &   oColumn.Comment & vbCrLf & vbCrLf
        Desc = Desc &   "h2. Raison et cas d'utilisation" & vbCrLf & vbCrLf
        Desc = Desc &   "..." & vbCrLf  & vbCrLf
        Desc = Desc &   "h2. Discussion" & vbCrLf & vbCrLf
        Desc = Desc &   "..." & vbCrLf  & vbCrLf
        Desc = Desc &   "h2. Contraintes " & vbCrLf & vbCrLf

        If Mid(oColumn.datatype,1,7) = "NUMERIC" Then
        If oColumn.LowValue <> oColumn.HighValue Then
            If Note<>"" Then Note = Note & " \\ \\ "
            Note = Note &  "Seules les valeurs entre *" & oColumn.LowValue & "* et *"  & oColumn.HighValue & "* sont permises."
            End if
        End If

        Note = ""
        If Not oColumn.Computed and oColumn.name<>"Key Section" and oColumn.name<>"Key Name" and oColumn.name<>"Key Value" Then
        If not oColumn.Domain is nothing Then
          If oColumn.Domain.ListOfValues <> "" Then
            If Note<>"" Then Note = Note & " \\ \\ "
            Note = CodeTab(oColumn.Domain.ListOfValues, oColumn.DefaultValue)
            End If
          End If
        End If

        Desc = Desc & Note & vbCrLf

        Desc = Desc &   "h2. Réferences et standard de données" & vbCrLf & vbCrLf
        Desc = Desc &   "..." & vbCrLf  & vbCrLf
        
        Desc = Desc &   "  {column}" & vbCrLf

        Desc = Desc &   "{section}" & vbCrLf

        Set File = oFileSystemObject.OpenTextFile(strPathWikiTable & "\" & lcase(oTable.Code) & "\" & lcase(oColumn.Code) & ".wiki", ForWriting, true)
        File.Write Desc & vbCrLf
        File.Close

        Wiki = Wiki & "tables" & "/" & lcase(oTable.Code) & "/" & lcase(oColumn.Code) & " """ & oColumn.Name & " (" & lcase(oTable.Stereotype) & ")" & """ """ & oTable.Name & """ " & StrSpace & vbCrLf
    
      End If
    Next

  End If
Next

Set File = oFileSystemObject.OpenTextFile(strPathWiki & "\wiki.properties", ForWriting, true)
File.Write Wiki & vbCrLf
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

Function YesNo(Flag)
    if Flag then
      YesNo = "!DEV:Images^yes.png|align=center!"
    else
      YesNo = "!DEV:Images^no.png|align=center!"
    end if
end Function

Function Style
  Style = "{style}" & vbCrLf
  Style = Style & "table.register{" & vbCrLf
  Style = Style & "  border:    1px solid #cccccc;" & vbCrLf
  Style = Style & "  border-collapse: collapse;}" & vbCrLf
  Style = Style & "table.register thead    { " & vbCrLf
  Style = Style & "  border-bottom: 1px solid #cccccc}" & vbCrLf
  Style = Style & "table.register thead td { " & vbCrLf
  Style = Style & "  border-bottom: 1px solid #cccccc;}" & vbCrLf
  Style = Style & "table.register td { border-bottom: 1px dotted #AAAAAA; vertical-align: top; padding: 2px;}" & vbCrLf
  Style = Style & "table.register tr.h1 td { " & vbCrLf
  Style = Style & "  border-bottom: 1px solid #cccccc; " & vbCrLf
  Style = Style & "  background-color: #f0f0f0;" & vbCrLf
  Style = Style & "  color: #003366;" & vbCrLf
  Style = Style & "  font-weight: bold;}" & vbCrLf
  Style = Style & "table.register tr.h2 td { " & vbCrLf
  Style = Style & "  border-bottom: 1px solid #cccccc; " & vbCrLf
  Style = Style & "  background-color: #f0f0f0;" & vbCrLf
  Style = Style & "  color: #003366;}" & vbCrLf
  Style = Style & "border: 1px solid #cccccc; background-color: #FAFAFA}" & vbCrLf
  Style = Style & "table.register tr.h3 td { border: 1px solid #003366}" & vbCrLf
  Style = Style & "{style}" & vbCrLf & vbCrLf
end Function

Function Codes(ListOfValues)
  Codes = Split(ListOfValues, vbNewLine, -1, 1)
end Function

Function CodeValue(Code)
  CodeValue = Split(ListOfValues, vbTab, -1, 1)
  CodeValue = CodeValue(0)
end Function

Function CodeDescription(Code)
  CodeDescription = Split(ListOfValues, vbTab, -1, 1)
  CodeValue = CodeValue(1)
End Function

Function CodeTab(ListOfValues, DefaultValue)
  S = "Seules les valeurs suivantes sont permises." & vbCrLf
  S = S &      "{table:class=register|width=50%}" & vbCrLf
  S = S &   "  {thead}" & vbCrLf
  S = s &   "    {tr:class=h1}" & vbCrLf
  S = S &   "      {td:rowspan=2}Valeur{td}" & vbCrLf
  S = S &   "      {td:colspan=2}Code{td}" & vbCrLf
  S = S &   "    {tr}" & vbCrLf
  S = S &   "    {tr:class=h2}" & vbCrLf
  S = S &   "      {td}Court{td}" & vbCrLf
  S = S &   "      {td}Long{td}" & vbCrLf
  S = S &   "    {tr}" & vbCrLf
  S = S &   "  {thead}" & vbCrLf
  S = S &   "  {tbody}" & vbCrLf
  Values = Split(ListOfValues, vbNewLine, -1, 1)
  For Ni=0 To UBound(Values)
    Value = Values(Ni)
    Value = Split(Value, vbTab, -1, 1)
    ValueText =  Split(Value(1), "|", -1, 1)
    If  Value(0)=DefaultValue then
      If UBound(ValueText)>0 Then 
        S = S &  "    {tr}{td}*" & Value(0) & "*{td}{td}*" & ValueText(0) & "*{td}{td}*" & ValueText(1)  & "*{td}{tr}" & vbCrLf
      Else
        S = S &  "    {tr}{td}*" & Value(0) & "*{td}{td}*" & ValueText(0) & "*{td}{td}{td}{tr}" & vbCrLf
      End If
    Else
      If UBound(ValueText)>0 Then 
        S = S &  "    {tr}{td}" & Value(0) & "{td}{td}" & ValueText(0) & "{td}{td}" & ValueText(1)  & "{td}{tr}" & vbCrLf
      Else
        S = S &  "    {tr}{td}" & Value(0) & "{td}{td}" & ValueText(0) & "{td}{td}{td}{tr}" & vbCrLf
      End If
    End If
  Next
  S = S &   "  {tbody}" & vbCrLf
  S = S &      "{table}" & vbCrLf
  CodeTab = S
End Function

Function CodeTabSimple(ListOfValues, DefaultValue)
  S = "Seules les valeurs suivantes sont permises." & vbCrLf
  Values = Split(ListOfValues, vbNewLine, -1, 1)
  For Ni=0 To UBound(Values)
    Value = Values(Ni)
    Value = Split(Value, vbTab, -1, 1)
    ValueText =  Split(Value(1), "|", -1, 1)
    If  Value(0)=DefaultValue then
      If UBound(ValueText)>1 Then 
        S = S &  "* *" & Value(0) & "*: " & ValueText(0) & " - " & ValueText(1)  & " (Valeur par défaut)" & vbCrLf
      Else
          S = S &  "* *" & Value(0) & "*: " & ValueText(0)  & " (Valeur par défaut)" & vbCrLf
      End If
      Else
      If UBound(ValueText)>1 Then 
        S = S &  "* *" & Value(0) & "*: " & ValueText(0) & " - " & ValueText(1) & vbCrLf
      Else
          S = S &  "* *" & Value(0) & "*: " & ValueText(0) & vbCrLf
      End If
    End If
  Next
  CodeTabSimple = S
End Function


Function Alias(S)
  If S="??" Then
    Alias="A définir"
  Else
    Alias=S
  End If
End Function
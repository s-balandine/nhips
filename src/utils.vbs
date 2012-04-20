' Constantes gloables
' .......................................

const ForReading   = 1 ' Open a File for reading only. You can't write to this oFile.
const ForWriting   = 2 ' Open a File for writing.
const ForAppending = 8 ' Open a File And write to the end of the oFile.

' Fonctions
' .......................................

Function CreateFolder(Folder)
	If Not oFileSystemObject.FolderExists(Folder) Then oFileSystemObject.CreateFolder(Folder)
	CreateFolder=Folder
End Function

function Sort(MyList)
  Dim sortFile, myArray, ts, i, j, temp, line, report
  myArray=Split(MyList,vbCrLf, -1, vbtextcompare)
  for i = UBound(myArray) - 1 To 0 Step -1
    for j= 0 to i
        if myArray(j)>myArray(j+1) Then
            temp=myArray(j+1)
            myArray(j+1)=myArray(j)
            myArray(j)=temp
        end if
    next
  next
  For Each line In myArray
    If Len(line) <> 0 Then
    report = report & line & vbcrlf
    End If
  Next
  Sort = report
end function

function SortCollection(Collection)
  Dim sortFile, Array, ts, i, j, temp, line, report, Names, ANames
  For Each Item In Collection
    Names=Names & Item.Name & vbCrLf
  Next
  ANames=Split(Names,vbCrLf, -1, vbtextcompare)
  For i = UBound(ANames) - 1 To 0 Step -1
    For j= 0 to i
        if ANames(j)>ANames(j+1) Then
            temp=ANames(j+1)
            ANames(j+1)=ANames(j)
            ANames(j)=temp
        end if
    Next
  Next
  For IdxTarget=1 To UBound(ANames)-10
    AName=ANames(IdxTarget)
    If Len(AName) <> 0 Then
      IdxSource = 0
      For Each Item In Collection
         If Item.Name=AName Then
           Collection.Move IdxTarget, IdxSource 
         End If
         IdxSource = IdxSource + 1
      Next
    End If
  Next
  set SortCollection = Collection
end function

function Code(MyList)
  Dim sortFile, myArray, ts, i, j, temp, line, report
  myArray=Split(MyList,vbCrLf, -1, vbtextcompare)
  For Each line In myArray
    If Len(line) <> 0 Then
      report = report & Mid(line,1,Instr(line,"¦")-1) & "=" & Mid(line,Instr(line,"¦")+1) & vbcrlf
    End If
  Next
  Code = report
end function

function Rank(MyList)
  Dim sortFile, myArray, ts, i, j, temp, line, report
  myArray=Split(MyList,vbCrLf, -1, vbtextcompare)
  For Each line In myArray
    If Len(line) <> 0 Then
      i = i + 1
      if i<10 then report = report & "S00" & i _
      else if i<100 then report = report & "S0" & i _
      else report = report & "S" & i      
      report = report & "_" & line & "¦" & vbcrlf
    End If
  Next
  Rank = report
end Function

Function writeUnicodeADODB(txtInput,filePath)
       
        ' Create and open stream
                Dim objStream
                Set objStream = CreateObject("ADODB.Stream")
                objStream.Open

        'Reset the position and indicate the charactor encoding
                objStream.Position = 0
                objStream.Charset = "UTF-8"

        'Write to the steam
                objStream.WriteText txtInput

        'Save the stream to a file
                filePath = filePath
                objStream.SaveToFile filePath, 2 ' overwrite if exists
       
        ' Return filepath with an @ so that imagemagick understands that it's a file
                writeUnicodeADODB = "@" & RemoveBOM(filePath)
       
        ' Kill stream
                Set objStream = Nothing
               
End Function

' Removes the Byte Order Mark - BOM from a text file with UTF-8 encoding
' The BOM defines that the file was stored with an UTF-8 encoding.
Public function RemoveBOM(filePath)
       
        ' Create a reader and a writer
                Dim writer,reader, fileSize
                Set writer = CreateObject("Adodb.Stream")
                Set reader = CreateObject("Adodb.Stream")
       
        ' Load from the text file we just wrote
                reader.Open
                reader.LoadFromFile filePath
       
        ' Copy all data from reader to writer, except the BOM
                writer.Mode=3
                writer.Type=1
                writer.Open
                reader.position=5
                reader.copyto writer,-1

        ' Overwrite file
                writer.SaveToFile filePath,2
       
        ' Return file name
                RemoveBOM = filePath

        ' Kill objects
                Set writer = Nothing
                Set reader = Nothing

end Function
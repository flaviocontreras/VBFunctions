CriaEstruturaDB(NomeDoServidor, NomeTabela, Usuario, Senha, IncluiBD)

'Cria uma string de conexao para um servidor de banco de dados SQL Server, dado Nome do servidor, nome da tabela, usuario e senha
Function CriaConnectionString(NomeDoServidor, NomeTabela, Usuario, Senha, IncluiBD)
	Dim CS
	CS = "Provider=SQLOLEDB.1; Data Source=" & NomeDoServidor
	If IncluiBD=True Then CS = CS & "; Initial Catalog=" & NomeTabela
	If Usuario<>"" Then
		CS = CS & ";User Id=" & Usuario & ";Password=" & Senha & ";"
	Else
		CS = CS & "; Integrated Security=SSPI"
	End If
	CriaConnectionString = CS 
End Function

Function CriaEstruturaDB(NomeDoServidor, NomeTabela, Usuario, Senha, IncluiBD)
On Error Resume Next
Dim objFso, objFolder, folder, file, objFile, text, objConn, cmd, cs

Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFso.GetFolder("\")
Set objConn = CreateObject("ADODB.Connection")
cs = CriaConnectionString(NomeDoServidor, NomeTabela, Usuario, Senha,True)

objConn.Open cs
objConn.BeginTrans
Set cmd = CreateObject("ADODB.Command")
Set cmd.ActiveConnection = objConn

For Each folder In objFolder.SubFolders
	For Each file In folder.Files
		Set objFile = objFso.OpenTextFile(file.Path, 1)
		text = objFile.ReadAll
	
		cmd.CommandText = text
		cmd.CommandType = 1
		cmd.Execute
		
		If Err.Number <> 0 Then
			objConn.RollbackTrans
			objConn.Close
			Set objConn = Nothing
			Exit Function
		End If

	Next
Next

objConn.CommitTrans
objConn.Close
Set objConn = Nothing
End Function
<%
Class Database
  Private pConnectionString
  Private pConnection

  Public Property Let ConnectionString
    ConnectionString = pConnectionString
  End Property

  Public Default Function construct(connectionString)
    pConnectionString = connectionString
    Set construct = Me
  End Function
  
  Private Sub Class_Terminate
    If Not pConnection Is Nothing Then 
      Call pConnection.Close()
    End If
    Set pConnection = Nothing
  End Sub
  
  Public Function ExecuteQuery(commandText, parameters)
    ' Turn me in to a recordset
  End Function
  
  Public Sub ExecuteNonQuery(commandText, parameters)
    ' Execute a non query
  End Sub
  
  Private Function GetOpenConnection()
    If pConnection Is Nothing Then
      Set pConnection = CreateObject("ADODB.Connection")
    End If
    
    If pConnection.State = 0 Then 'Closed
      Call pConnection.Open(pConnectionString)
    End If
    
    Set GetOpenConnection = pConnection
  End Function
End Class
%>

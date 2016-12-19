<%
Class Database
  Private pConnectionString
  Private pConnection

  Public Property Get ConnectionString
    ConnectionString = pConnectionString
  End Property

  Public Default Function construct(connectionString)
    pConnectionString = connectionString
    Set pConnection = CreateObject("ADODB.Connection")
    Set construct = Me
    Call pConnection.Open(connectionString)
  End Function
  
  Private Sub Class_Terminate
    Call pConnection.Close()
    Set pConnection = Nothing
  End Sub
  
  Public Function ExecuteQuery(commandText, parameters)
    Dim rs : Set rs = CreateObject("ADODB.Recordset")
    
    rs.Open commandText, pConnection, adOpenStatic, adLockOptimistic
    
    Set ExecuteQuery = rs
    Set rs = Nothing
  End Function
  
  Public Function ExecuteNonQuery(commandText, parameters)
    Set ExecuteNonQuery = pConnection.Execute(commandText)
  End Function
End Class
%>

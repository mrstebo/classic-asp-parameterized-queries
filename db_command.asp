<%
Class DbCommand
  Private pCommandText
  Private pConnection
  Private pParameters

  Public Default Function construct(commandText, connection)
    pCommandText = commandText
    Set pConnection = connection
    pParameters = Array()
    Set construct = Me
  End Function
  
  Public Sub AddParameter(parameter)
    ReDim Preserve pParameters(UBound(parameters) + 1)
    Set parameters(UBound(parameters)) = parameter
  End Sub
  
  Public Sub AddParameters(parameters)
    For Each parameter in parameters
      Call AddParameter(parameter)
    Next
  End Sub
End Class
%>

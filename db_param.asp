<%
Class DbParam
  Public ParameterName
  Public Value
  
  Public Default Function construct(parameterName, value)
    ParameterName = parameterName
    Value = value
    Set construct = Me
  End Function
End Class
%>

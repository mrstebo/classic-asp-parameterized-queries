<!--#include FILE='database.asp'-->
<%
Dim filename : filename = "test.mdb"
Dim provider : provider = "Microsoft.Jet.OLEDB.4.0"
'Dim provider : provider = "Microsoft.ACE.OLEDB.12.0"
Dim connectionString : connectionString = "Data Source=" & filename & ";Provider=" & provider & ";"
Dim db : Set db = (New Database)(connectionString)
Dim rs : Set rs = db.ExecuteQuery("SELECT * FROM [products]", Nothing)
%>

<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>classic-asp-parameterized-queries</title>
</head>
<body>
  <h1>Test Page</h1>
  <h3>Getting rows from the <strong>products</strong> table</h3>
  <table>
    <thead>
      <tr>
        <th>ID</th>
        <th>Name</th>
        <th>Description</th>
        <th>Price</th>
        <th>Number In Stock</th>
      </tr>
    </thead>
    <tbody>
      <% While Not rs.EOF %>
        <tr>
          <td>
            <% rs("ID") %>
          </td>
          <td>
            <% rs("ProductName") %>
          </td>
          <td>
            <% rs("Description") %>
          </td>
          <td>
            <% rs("Price") %>
          </td>
          <td>
            <% rs("NumberInStock") %>
          </td>
        </tr>
        <% rs.MoveNext %>
      <% WEnd %>
    </tbody>
  </table>
</body>
</html>

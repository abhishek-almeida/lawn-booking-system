Attribute VB_Name = "db_connect"
Public cn As adodb.Connection
Public rs As adodb.Recordset

Sub main()
    Set cn = New adodb.Connection
    cn.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=localhost;Database=lawn_booking_db;User=root;Password=based;Option=3;"
    cn.Open
    admin_login.Show
End Sub



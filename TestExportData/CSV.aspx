<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CSV.aspx.cs" Inherits="TestExportData.CSV" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <input type="button" runat="server" OnServerClick="OnServerClick" value="测试直接传入DataTable" ID="btnTest"/>
        <input type="button" runat="server" OnServerClick="btnTest2_OnServerClick" value="测试直接传入DataTable以及标题数组" ID="btnTest2"/>
        <input type="button" runat="server" OnServerClick="btnTest3_OnServerClick" value="测试直接传入DataTable以及标题数组以及字段名称数组" ID="btnTest3"/>
        <br/>
        <asp:Button runat="server" ID="testBtnJS" OnClick="testBtnJS_OnClick"/>
    </div>
    </form>
    <script type="text/javascript">
        <%=jsmsg%>
    </script>
</body>
</html>

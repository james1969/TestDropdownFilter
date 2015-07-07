<!--#include virtual="../wheelsonline.ca/_millbrook_domain_specific/15_connection.dock"-->


<script>

function enable() {
    document.getElementById("model").disabled=false;
    var varMakeSelected = document.getElementById("make").value;
    alert(varMakeSelected);
}

</script>


<h5>Inventory search can be made by Year, Make, Model or any combination of these.<br>A minimum of a Make or a Model must be selected</h5>

<div style="border:#666666 1px solid; width:600px; margin:auto; margin-bottom:20px; background-color:#CCCCCC" align="center">
<form action="inventory.asp?searchinventory=<%=varAction%>" method="post" style="margin:0px">

<table width="600" border="0" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td>
    <select name="year" id="year">
    <option value="">Year</option>
        <%sqlYear = "SELECT DISTINCT vYear FROM wo_carinfo WHERE dealercode = 'p4l' AND status = 'Available' ORDER BY vYear DESC"
        set rsYear = conn.execute(sqlYear)%>
        <%do until rsYear.eof
                if rsYear("vYear") = varYear then%>
        		<option value="<%=rsYear("vYear")%>" selected><%=rsYear("vYear")%></option>
        		<%else %>
        		<option value="<%=rsYear("vYear")%>"><%=rsYear("vYear")%></option>
        		<%end if
                rsYear.movenext
                loop
                rsYear.Close
                Set rsYear = Nothing%>
    </select>
    </td>

    <td>
    <select name="make" id="make" onchange="enable()">
    <option value="">Make</option>
        <%sqlMake = "SELECT DISTINCT Make FROM wo_carinfo WHERE dealercode = 'p4l' AND status = 'Available' ORDER BY Make"
        set rsMake = conn.execute(sqlMake)%>
        <%do until rsMake.eof
                if rsMake("Make") = varMake then%>
        		<option value="<%=rsMake("Make")%>" selected><%=rsMake("Make")%></option>
        		<%else %>
        		<option value="<%=rsMake("Make")%>"><%=rsMake("Make")%></option>
        		<%end if
                rsMake.movenext
                loop
                rsMake.Close
                Set rsMake = Nothing%>
    </select>
    </td>

    <td>
    <select disabled name="model" id="model">
    <option value="">Model</option>
        <%sqlModel = "SELECT DISTINCT Model FROM wo_carinfo WHERE dealercode = 'p4l' AND make = '" & varMakeSelected &  "' AND Model <> '' AND status = 'Available' ORDER BY Model"
        set rsModel = conn.execute(sqlModel)%>
        <%do until rsModel.eof
                if rsModel("Model") = varMake then%>
        		<option value="<%=rsModel("Model")%>" selected><%=rsModel("Model")%></option>
        		<%else %>
        		<option value="<%=rsModel("Model")%>"><%=rsModel("Model")%></option>
        		<%end if
                rsModel.movenext
                loop
                rsModel.Close
                Set rsModel = Nothing%>
    </select>
    </td>

    <td>
    <input type="submit" name="searchinventory" id="button" value="Search Inventory"/>
    </td>

  </tr>
</table>

</form>
</div>










<div align="center">

## Fill a DataReport with a Recordset


</div>

### Description

This code enable you fill a DataReport with the information contained in a Recordset. In this example, I have a connection with a database and I fill the Recordset with a SQL sentence.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-01-27 15:14:42
**By**             |[Mauricio Prieto Garc&\#237;a](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mauricio-prieto-garc-237-a.md)
**Level**          |Beginner
**User Rating**    |4.9 (44 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Fill\_a\_Dat1535411272003\.zip](https://github.com/Planet-Source-Code/mauricio-prieto-garc-237-a-fill-a-datareport-with-a-recordset__1-41383/archive/master.zip)





### Source Code

<pre>
This code enable you fill a DataReport with the
information contained in a Recordset. In this
example, I have a connection with a database and
I fill the Recordset with a SQL sentence.
<b>First</b>, you have to add a DataReport to your project:
	Project > Add DataReport
<b>Design your report.</b>
The RptLabel tools are usually used as headers,
but you can also use them in the report’s body
and another places. The RptTextBox can be only
on the body of the report and this tool receive
the information from de recordset. In Properties
you must give a different name to each
RptTextBox.
The item’s parameter is the name of the
RptTextBox which will receive the information.
<b>
DataReport1.Sections(“[SectionName”).Controls.Item(“[RptTextBoxName]”).DataField
</b>
I put the next code in the click event of a menu
but you can put it in any other control
(i.e.,CommandButton)
Dim Rs As New ADODB.Recordset
 Set Rs = Cn.Execute("Select ClientPK,NameClient,RepC,DomClient,TelClient,
CelClient,RFCClient,DateClient from Client order
by(NomClient)")
 With DataReport1
 .DataMember = vbNullString
 Set .DataSource = Rs
 .Caption = "This is the Title of the DataReport window"
 With .Sections("Section1").Controls
 .Item("tNo").DataField = Rs.Fields(0).Name
 .Item("tName").DataField = Rs.Fields(1).Name
 .Item("tRep").DataField = Rs.Fields(2).Name
 .Item("tDom").DataField = Rs.Fields(3).Name
 .Item("tTel").DataField = Rs.Fields(4).Name
 .Item("tCel").DataField = Rs.Fields(5).Name
 .Item("tRFC").DataField = Rs.Fields(6).Name
 .Item("tDate").DataField = Rs.Fields(7).Name
 End With
 .Show
 End With
<br>
I hope it will be useful to you
</pre>


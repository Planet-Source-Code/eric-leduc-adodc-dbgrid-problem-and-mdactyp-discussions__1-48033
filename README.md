<div align="center">

## Adodc \+ dbGrid problem and MdacTyp discussions


</div>

### Description

I had problems using the datagrid attached to an ADOdc control.

Here is what I found.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Eric Leduc](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eric-leduc.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eric-leduc-adodc-dbgrid-problem-and-mdactyp-discussions__1-48033/archive/master.zip)





### Source Code

When using an ADOdc to connect to a Access database, where the command is set to adCmdUnknown to be able to set the recordSource using a SQL statement, and attach a datagrid to that control, if you attempt to add a new record to the empty list(corresponding to an empty table) you will get the error message 'current row unavailable' <BR><BR>
This problem is not documented in Microsoft knowledgebase, or if I recall, the workaround they suggest does not work. <BR><BR>
I found in some newsgroup that if you want to change the recordSource of the AdoDc , to avoid the problem with the grid, you first have to disconnect the grid from the Ado control <BR><code>
Set dataGrid.DataSource= Nothing<BR>
'then change the ado query<BR>
AdoDc.RecordSource = "SELECT * from myTable WHERE myField = 'someStringValueforExample'"<BR>
AdoDc.Refresh<BR>
Set dataGrid.DataSource = AdoDc </code><BR><BR>
if the table is empty, and you use the datagrid to add a new record, you will not get the current row error using this technique. <BR><BR>
Other matters:<BR>
You will realize that in order to use the vb data wizard to create your startup data forms (wich is good), you will need to reference the msADO 2.7 Library <BR><BR>
If you do so, to redistribute your application properly, you have to download the appropriate (2.7) MDAC_Type.exe from microsoft and put this file in the <BOLD> C:\Program Files\Microsoft Visual Studio\VB98\Wizards\PDWizard\Redist </BOLD> folder.<BR><BR>
Hope this will help someone.... I sure would have appreciated it when I had problems.


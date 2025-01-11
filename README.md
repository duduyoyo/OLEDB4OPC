# OLEDB4OPC
Introduce the lowest level DB programming in Classic OPC for performance reason!

Microsoft has revived OLE DB in recent years with new release of OLE DB driver. The new features like encryption in this new driver make it an ideal candidate of database programming in Cloud era, especially with performance in mind. But, due to its long backward history, it is hard to find a good example to apply it with Classic OPC. This solution will fill the gap by showing how to use OLE DB programming in Classic OPC DA server, and insert/select data to/from database quickly. OPC data callback feature is used in this example.

<h2>Pre-requiste</h2>
1. Download the latest OLE DB driver from Microsoft <a href="https://learn.microsoft.com/en-us/sql/connect/oledb/download-oledb-driver-for-sql-server?view=sql-server-ver16">site</a>.<p></p>
2. In SQL Server create a TEST database, and then create a table by running following query,<p></p>
use TEST;<p></p>
CREATE TABLE [dbo].[OPCDA](<p></p>
[Tag] [nvarchar](50) NOT NULL,<p></p>
[Value] [real] NOT NULL,<p></p>
[Time] [datetime] NOT NULL,<p></p>
[Quality] [tinyint] NOT NULL<p></p>
);<p></p>
3. Double click OPC.udl to establish a database connection to SQL Server. Only leave minimum attributes to satisfy backward compatibility need as already shown in this file.<p></p>
4. Install an OPC DA server in the same box.<p></p>
5. To run as a 64-bit application you need register 64-bit proxy dlls first from command line, like "C:\Windows\System32\regsvr32 x64Proxy\opccomn_ps.dll" and "C:\Windows\System32\regsvr32 x64Proxy\opcproxy.dll"<p></p>

<h2>Compile and run</h2>
This console project is compiled and run under 64-bit debug mode for Visual Studio 2022.<p></p>

<h2>Console output</h2>
<img src="https://github.com/duduyoyo/OLEDB4OPC/assets/13662339/3908cc99-adbb-444f-8174-30309b14b08e" width=70%>

<h2>Related contribution</h2>
<a href="https://github.com/duduyoyo/WebSocket4OPC">WebSocket4OPC</a>, a modern approach to access plant data anywhere and any way!<p>
<a href="https://github.com/duduyoyo/WebSocket4Fragment">WebSocket4Fragment</a>, a unique way to handle WebSocket fragment explicitly in run time!

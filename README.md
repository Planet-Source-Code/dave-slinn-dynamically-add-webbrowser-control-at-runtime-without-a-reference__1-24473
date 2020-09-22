<div align="center">

## Dynamically Add WebBrowser Control at runtime without a Reference

<img src="PIC20016262137211102.jpg">
</div>

### Description

Allows VB applications to determine at run-time if Internet Explorer (4.0 or later) is installed, and if so, creates a WebBrowser. If not, a trappable error allows program to continue.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dave Slinn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dave-slinn.md)
**Level**          |Beginner
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dave-slinn-dynamically-add-webbrowser-control-at-runtime-without-a-reference__1-24473/archive/master.zip)





### Source Code

<font face="Tahoma" size="2"><p>Add a menu item named 'mnuCreate' with a caption of "&Create
WebBrowser"</p>
<p>Place the following code into a standard VB 6.0 form.</p></font>
<hr>
<p><font face="Courier New" size="2"><br>
<font color="#0000FF">Private</font> m_WebControl <font color="#0000FF"> As</font> VBControlExtender<br>
<br>
<font color="#0000FF">Private Sub</font> Form_Resize()<br>
On Error Resume Next<br>
<font color="#008000">   </font> <font color="#008000">' resize webbrowser to entire size of form</font><br>
    m_WebControl.Move 0, 0, ScaleWidth, ScaleHeight<br>
<font color="#0000FF">End Sub</font><br>
<br>
<font color="#0000FF">Private Sub</font> mnuCreate_Click()<br>
<font color="#0000FF">On Error GoTo</font> ErrHandler<br>
<br>
<font color="#008000">   </font> <font color="#008000">' attempting to add WebBrowser here ('Shell.Explorer.2' is registered<br>
    ' with Windows if a recent (>= 4.0) version of Internet Explorer is installed<br>
</font><font color="#0000FF">   </font><font color="#008000"> </font><font color="#0000FF">Set</font> m_WebControl = Controls.Add("Shell.Explorer.2", "webctl", Me)<br>
<br>
<font color="#008000">   </font> <font color="#008000">' if we got to here, there was no problem creating the WebBrowser<br>
    ' so we should size it properly and ensure it's visible<br>
</font>   <font color="#008000"> </font>m_WebControl.Move 0, 0, ScaleWidth, ScaleHeight<br>
    m_WebControl.Visible = <font color="#0000FF"> True</font><br>
<br>
<font color="#008000">   </font> <font color="#008000">' use the Navigate method of the WebBrowser control to open a<br>
    ' web page<br>
</font>   <font color="#008000"> </font>m_WebControl.object.navigate "http://www.planet-source-code.com"<br>
<br>
<font color="#0000FF">   </font> <font color="#0000FF">Exit Sub</font><br>
ErrHandler:<br>
    MsgBox "Could not create WebBrowser control", vbInformation<br>
<font color="#0000FF">End Sub</font></font><br>
</p>


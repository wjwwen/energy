Option Explicit

Sub InputTextCrude()
    Dim oPPApp As Object, oPPPrsn As Object, oPPSlide As Object
    Dim oPPShape As Object
    Dim FlName As String

    ' Read PPT file name
    FlName = "C:\Users\jingwen.wang\IHS Markit\Downstream_SG - Documents\Jobs\Multi Client Studies\OPSIS-Product Architecture\Crude Oil Module\2022\Sep2022\Asia and Middle East Crude Oil Markets Short-Term Outlook - Sep 2022.PPTX"

    ' Establish PowerPoint application object
    On Error Resume Next
    Set oPPApp = GetObject(, "PowerPoint.Application")

    If Err.Number <> 0 Then
        Set oPPApp = CreateObject("PowerPoint.Application")
    End If
    Err.Clear
    On Error GoTo 0

    oPPApp.Visible = True

    ' Open PPT file
    Set oPPPrsn = oPPApp.Presentations.Open(FlName)
    
    ' ----- CHINA CRUDE TRADE
    ' Slide with shape to input text
    Set oPPSlide = oPPPrsn.Slides(23)

    ' Shape name
    ' ----- IRAQ-CHINA ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("IraqChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B3").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C3").Value

    Set oPPShape = oPPSlide.Shapes("IraqChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D3").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqChina4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E3").Value

    ' ----- IRAN-CHINA ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("IranChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B4").Value
    
    Set oPPShape = oPPSlide.Shapes("IranChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C4").Value

    Set oPPShape = oPPSlide.Shapes("IranChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D4").Value
    
    Set oPPShape = oPPSlide.Shapes("IranChina4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E4").Value

    ' ----- KUWAIT-CHINA ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("KuwaitChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B5").Value
    
    Set oPPShape = oPPSlide.Shapes("KuwaitChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C5").Value

    Set oPPShape = oPPSlide.Shapes("KuwaitChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D5").Value
    
    Set oPPShape = oPPSlide.Shapes("KuwaitChina4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E5").Value

    ' ----- UAE-China ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("UAEChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B6").Value
    
    Set oPPShape = oPPSlide.Shapes("UAEChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C6").Value

    Set oPPShape = oPPSlide.Shapes("UAEChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D6").Value
    
    Set oPPShape = oPPSlide.Shapes("UAEChina4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E6").Value

    '----- Saudi-China ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("SaudiChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B7").Value
    
    Set oPPShape = oPPSlide.Shapes("SaudiChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C7").Value

    Set oPPShape = oPPSlide.Shapes("SaudiChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D7").Value
    
    Set oPPShape = oPPSlide.Shapes("SaudiChina4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E7").Value

    '----- US-China ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("USChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B8").Value
    
    Set oPPShape = oPPSlide.Shapes("USChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C8").Value

    Set oPPShape = oPPSlide.Shapes("USChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D8").Value
    
    Set oPPShape = oPPSlide.Shapes("USChina4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E8").Value

    ' ----- Russia-China ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("RussiaChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B9").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C9").Value

    Set oPPShape = oPPSlide.Shapes("RussiaChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D9").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaChina4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E9").Value

    ' ----- Africa-China ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("AfricaChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B11").Value
    
    Set oPPShape = oPPSlide.Shapes("AfricaChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C11").Value

    Set oPPShape = oPPSlide.Shapes("AfricaChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D11").Value
    
    Set oPPShape = oPPSlide.Shapes("AfricaChina4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E11").Value

    ' ----- SOUTH KOREA CRUDE TRADE
    ' Slide with shape to input text
    Set oPPSlide = oPPPrsn.Slides(24)

    ' Shape name
    ' ----- IRAQ-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("IraqKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B17").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C17").Value

    Set oPPShape = oPPSlide.Shapes("IraqKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D17").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqKorea4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E17").Value

    ' ----- IRAN-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("IranKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B18").Value
    
    Set oPPShape = oPPSlide.Shapes("IranKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C18").Value

    Set oPPShape = oPPSlide.Shapes("IranKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D18").Value
    
    Set oPPShape = oPPSlide.Shapes("IranKorea4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E18").Value

    ' ----- KUWAIT-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("KuwaitKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B19").Value
    
    Set oPPShape = oPPSlide.Shapes("KuwaitKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C19").Value

    Set oPPShape = oPPSlide.Shapes("KuwaitKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D19").Value
    
    Set oPPShape = oPPSlide.Shapes("KuwaitKorea4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E19").Value

    ' ----- Qatar-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("QatarKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B20").Value
    
    Set oPPShape = oPPSlide.Shapes("QatarKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C20").Value

    Set oPPShape = oPPSlide.Shapes("QatarKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D206").Value
    
    Set oPPShape = oPPSlide.Shapes("QatarKorea4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E20").Value

    ' ----- UAE-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("UAEKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B21").Value
    
    Set oPPShape = oPPSlide.Shapes("UAEKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C21").Value

    Set oPPShape = oPPSlide.Shapes("UAEKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D21").Value
    
    Set oPPShape = oPPSlide.Shapes("UAEKorea4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E21").Value

    '----- Saudi-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("SaudiKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B22").Value
    
    Set oPPShape = oPPSlide.Shapes("SaudiKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C22").Value

    Set oPPShape = oPPSlide.Shapes("SaudiKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D22").Value
    
    Set oPPShape = oPPSlide.Shapes("SaudiKorea4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E22").Value

    '----- US-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("USKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B23").Value
    
    Set oPPShape = oPPSlide.Shapes("USKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C23").Value

    Set oPPShape = oPPSlide.Shapes("USKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D23").Value
    
    Set oPPShape = oPPSlide.Shapes("USKorea4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E23").Value

    ' ----- Russia-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("RussiaKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B24").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C24").Value

    Set oPPShape = oPPSlide.Shapes("RussiaKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D24").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaKorea4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E24").Value

    ' ----- Africa-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("AfricaKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B26").Value
    
    Set oPPShape = oPPSlide.Shapes("AfricaKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C26").Value

    Set oPPShape = oPPSlide.Shapes("AfricaKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D26").Value
    
    Set oPPShape = oPPSlide.Shapes("AfricaKorea4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E26").Value

    ' ----- JAPAN CRUDE TRADE
    ' Slide with shape to input text
    Set oPPSlide = oPPPrsn.Slides(25)

    ' Shape name
    ' ----- IRAQ-Japan ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("IraqJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B32").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqJapan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C32").Value

    Set oPPShape = oPPSlide.Shapes("IraqJapan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D32").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqJapan4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E32").Value

    ' ----- IRAN-Japan ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("IranJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B33").Value
    
    Set oPPShape = oPPSlide.Shapes("IranJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C33").Value

    Set oPPShape = oPPSlide.Shapes("IranJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D33").Value
    
    Set oPPShape = oPPSlide.Shapes("IranJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E33").Value

    ' ----- KUWAIT-Japan ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("KuwaitJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B34").Value
    
    Set oPPShape = oPPSlide.Shapes("KuwaitJapan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C34").Value

    Set oPPShape = oPPSlide.Shapes("KuwaitJapan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D34").Value
    
    Set oPPShape = oPPSlide.Shapes("KuwaitJapan4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E34").Value

    ' ----- Qatar-Japan ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("QatarJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B35").Value
    
    Set oPPShape = oPPSlide.Shapes("QatarJapan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C35").Value

    Set oPPShape = oPPSlide.Shapes("QatarJapan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D35").Value
    
    Set oPPShape = oPPSlide.Shapes("QatarJapan4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E35").Value

    ' ----- UAE-Japan ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("UAEJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B36").Value
    
    Set oPPShape = oPPSlide.Shapes("UAEJapan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C36").Value

    Set oPPShape = oPPSlide.Shapes("UAEJapan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D36").Value
    
    Set oPPShape = oPPSlide.Shapes("UAEJapan4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E36").Value

    '----- Saudi-Korea ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("SaudiJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B37").Value
    
    Set oPPShape = oPPSlide.Shapes("SaudiJapan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C37").Value

    Set oPPShape = oPPSlide.Shapes("SaudiJapan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D37").Value
    
    Set oPPShape = oPPSlide.Shapes("SaudiJapan4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E37").Value

    '----- US-Japan ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("USJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B38").Value
    
    Set oPPShape = oPPSlide.Shapes("USJapan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C38").Value

    Set oPPShape = oPPSlide.Shapes("USJapan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D38").Value
    
    Set oPPShape = oPPSlide.Shapes("USJapan4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E38").Value

    ' ----- Russia-Japan ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("RussiaJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B39").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaJapan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C39").Value

    Set oPPShape = oPPSlide.Shapes("RussiaJapan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D39").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaJapan4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E39").Value

    ' ----- INDIA CRUDE TRADE
    ' Slide with shape to input text
    Set oPPSlide = oPPPrsn.Slides(26)

    ' Shape name
    ' ----- Iran-India ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("IraqIndia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B32").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqIndia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C32").Value

    Set oPPShape = oPPSlide.Shapes("IraqIndia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D32").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqIndia4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E32").Value

    ' ----- Iraq-India ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("IraqIndia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B48").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqIndia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C48").Value

    Set oPPShape = oPPSlide.Shapes("IraqIndia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D48").Value
    
    Set oPPShape = oPPSlide.Shapes("IraqIndia4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E48").Value

    ' ----- KUWAIT-India ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("KuwaitIndia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B49").Value
    
    Set oPPShape = oPPSlide.Shapes("KuwaitIndia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C49").Value

    Set oPPShape = oPPSlide.Shapes("KuwaitIndia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D49").Value
    
    Set oPPShape = oPPSlide.Shapes("KuwaitIndia4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E49").Value

    ' ----- Qatar-India ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("QatarIndia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E50").Value
    
    Set oPPShape = oPPSlide.Shapes("QatarIndia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C50").Value

    Set oPPShape = oPPSlide.Shapes("QatarIndia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D50").Value
    
    Set oPPShape = oPPSlide.Shapes("QatarIndia4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E50").Value

    ' ----- UAE-India ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("UAEIndia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B51").Value
    
    Set oPPShape = oPPSlide.Shapes("UAEIndia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C51").Value

    Set oPPShape = oPPSlide.Shapes("UAEIndia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D51").Value
    
    Set oPPShape = oPPSlide.Shapes("UAEIndia4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E51").Value

    '----- Saudi-India ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("SaudiIndia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B52").Value
    
    Set oPPShape = oPPSlide.Shapes("SaudiIndia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C52").Value

    Set oPPShape = oPPSlide.Shapes("SaudiIndia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D52").Value
    
    Set oPPShape = oPPSlide.Shapes("SaudiIndia4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E52").Value

    '----- US-India ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("USIndia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B53").Value
    
    Set oPPShape = oPPSlide.Shapes("USIndia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C53").Value

    Set oPPShape = oPPSlide.Shapes("USIndia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D53").Value
    
    Set oPPShape = oPPSlide.Shapes("USIndia4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E53").Value

    ' ----- Russia-India ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("RussiaIndia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B54").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaIndia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C54").Value

    Set oPPShape = oPPSlide.Shapes("RussiaIndia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D54").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaIndia4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E54").Value

    ' ----- Africa-India ----- ----- ----- -----
    Set oPPShape = oPPSlide.Shapes("AfricaIndia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("B56").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaIndia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("C56").Value

    Set oPPShape = oPPSlide.Shapes("RussiaIndia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("D56").Value
    
    Set oPPShape = oPPSlide.Shapes("RussiaIndia4")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Linksout-Macro").Range("E56").Value


End Sub






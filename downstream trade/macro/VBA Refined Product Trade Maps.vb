Option Explicit

Sub InputText()
    Dim oPPApp As Object, oPPPrsn As Object, oPPSlide As Object
    Dim oPPShape As Object
    Dim FlName As String

    ' Read PPT file name
    FlName = "C:\Users\jingwen.wang\Desktop\Asia and Middle East RM STO Aug 2022_Kendrick Wee_V4.PPTX"

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
    
    ' ----- GASOLINE TRADE
    ' Slide with shape to input text
    Set oPPSlide = oPPPrsn.Slides(23)

    ' Shape name
    Set oPPShape = oPPSlide.Shapes("IndiaSG1")
    ' Write to shape
    '------------------- INDIA -------------------
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K12").Value
    
    Set oPPShape = oPPSlide.Shapes("IndiaSG2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L12").Value

    Set oPPShape = oPPSlide.Shapes("IndiaSG3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M12").Value
    
    Set oPPShape = oPPSlide.Shapes("IndiaAfrica1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K14").Value

    Set oPPShape = oPPSlide.Shapes("IndiaAfrica2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L14").Value
    
    Set oPPShape = oPPSlide.Shapes("IndiaAfrica3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M14").Value
    
    Set oPPShape = oPPSlide.Shapes("IndiaME1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K13").Value

    Set oPPShape = oPPSlide.Shapes("IndiaME2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L13").Value
    
    Set oPPShape = oPPSlide.Shapes("IndiaME3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M13").Value
    
    '------------------- CHINA -------------------
    Set oPPShape = oPPSlide.Shapes("ChinaSG1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K7").Value

    Set oPPShape = oPPSlide.Shapes("ChinaSG2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L7").Value
    
    Set oPPShape = oPPSlide.Shapes("ChinaSG3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M7").Value
    
    Set oPPShape = oPPSlide.Shapes("ChinaMsia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K6").Value

    Set oPPShape = oPPSlide.Shapes("ChinaMsia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L6").Value
    
    Set oPPShape = oPPSlide.Shapes("ChinaMsia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M6").Value
    
    '------------------- KOREA -------------------
    Set oPPShape = oPPSlide.Shapes("KoreaVn1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K19").Value

    Set oPPShape = oPPSlide.Shapes("KoreaVn2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L19").Value
    
    Set oPPShape = oPPSlide.Shapes("KoreaVn3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M19").Value

    Set oPPShape = oPPSlide.Shapes("KoreaSG1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K21").Value

    Set oPPShape = oPPSlide.Shapes("KoreaSG2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L21").Value
    
    Set oPPShape = oPPSlide.Shapes("KoreaSG3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M21").Value
    
    Set oPPShape = oPPSlide.Shapes("KoreaAus1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K22").Value

    Set oPPShape = oPPSlide.Shapes("KoreaAus2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L22").Value
    
    Set oPPShape = oPPSlide.Shapes("KoreaAus3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M22").Value
    
    Set oPPShape = oPPSlide.Shapes("KoreaJap1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K20").Value

    Set oPPShape = oPPSlide.Shapes("KoreaJap2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L20").Value
    
    Set oPPShape = oPPSlide.Shapes("KoreaJap3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M20").Value
    
    '------------------- SINGAPORE -------------------
    Set oPPShape = oPPSlide.Shapes("SGIndon1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K31").Value

    Set oPPShape = oPPSlide.Shapes("SGIndon2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L31").Value
    
    Set oPPShape = oPPSlide.Shapes("SGIndon3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M31").Value
    
    Set oPPShape = oPPSlide.Shapes("SGAus1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K36").Value

    Set oPPShape = oPPSlide.Shapes("SGAus2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L36").Value
    
    Set oPPShape = oPPSlide.Shapes("SGAus3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M36").Value

    Set oPPShape = oPPSlide.Shapes("SGMsia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K26").Value

    Set oPPShape = oPPSlide.Shapes("SGMsia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L26").Value
    
    Set oPPShape = oPPSlide.Shapes("SGMsia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M26").Value

    ' ----- JET/KEROSENE TRADE
    ' Slide with shape to input text
    Set oPPSlide = oPPPrsn.Slides(27)

    '------------------- CHINA -------------------
    Set oPPShape = oPPSlide.Shapes("Jet-ChinaEur1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K82").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-ChinaEur2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L82").Value

    Set oPPShape = oPPSlide.Shapes("Jet-ChinaEur3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M82").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-ChinaNAM1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K83").Value

    Set oPPShape = oPPSlide.Shapes("Jet-ChinaNAM2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L83").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-ChinaNAM3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M83").Value
    
    '------------------- INDIA -------------------
    Set oPPShape = oPPSlide.Shapes("Jet-IndiaEur1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K87").Value

    Set oPPShape = oPPSlide.Shapes("Jet-IndiaEur2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L87").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-IndiaEur3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M87").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-IndiaAfrica1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K88").Value

    Set oPPShape = oPPSlide.Shapes("Jet-IndiaAfrica2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L88").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-IndiaAfrica3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M88").Value
    
    '------------------- SINGAPORE -------------------
    Set oPPShape = oPPSlide.Shapes("Jet-SGVn1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K94").Value

    Set oPPShape = oPPSlide.Shapes("Jet-SGVn2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L94").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-SGVn3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M94").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-SGMsia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K93").Value

    Set oPPShape = oPPSlide.Shapes("Jet-SGMsia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L93").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-SGMsia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M93").Value

    '------------------- KOREA -------------------
    Set oPPShape = oPPSlide.Shapes("Jet-KorJap1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K101").Value

    Set oPPShape = oPPSlide.Shapes("Jet-KorJap2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L101").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-KorJap3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M101").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-KorChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K102").Value

    Set oPPShape = oPPSlide.Shapes("Jet-KorChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L102").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-KorChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M102").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-KorAus1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K104").Value

    Set oPPShape = oPPSlide.Shapes("Jet-KorAus2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L104").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-KorAus3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M104").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-KorNAM1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K105").Value

    Set oPPShape = oPPSlide.Shapes("Jet-KorNAM2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L105").Value
    
    Set oPPShape = oPPSlide.Shapes("Jet-KorNAM3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M105").Value

    ' ----- DIESEL TRADE
    ' Slide with shape to input text
    Set oPPSlide = oPPPrsn.Slides(31)

    '------------------- CHINA -------------------
    Set oPPShape = oPPSlide.Shapes("Diesel-ChinaPhil1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K44").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-ChinaPhil2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L44").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-ChinaPhil3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M44").Value
    
    '------------------- INDIA -------------------
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaSG1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K51").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaSG2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L51").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaSG3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M51").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaEur1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K53").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaEur2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L53").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaEur3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M53").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaAus1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K54").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaAus2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L54").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaAus3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M54").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaAfrica1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K55").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaAfrica2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L55").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaAfrica3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M55").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaME1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K56").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaME2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L56").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-IndiaME3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M56").Value

    '------------------- SINGAPORE -------------------
    Set oPPShape = oPPSlide.Shapes("Diesel-SGMSia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K62").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-SGMSia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L62").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-SGMSia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M62").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-SGAus1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K65").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-SGAus2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L65").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-SGAus3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M65").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-SGIndon1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K61").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-SGIndon2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L61").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-SGIndon3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M61").Value
    
    '------------------- KOREA -------------------
    Set oPPShape = oPPSlide.Shapes("Diesel-KorSG1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K72").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-KorSG2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L72").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-KorSG3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M72").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-KorChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K74").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-KorChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L74").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-KorChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M74").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-KorAus1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K75").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-KorAus2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L75").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-KorAus3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M75").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-KorPhil1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K76").Value

    Set oPPShape = oPPSlide.Shapes("Diesel-KorPhil2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L76").Value
    
    Set oPPShape = oPPSlide.Shapes("Diesel-KorPhil3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M76").Value

    ' ----- FUEL OIL TRADE
    ' Slide with shape to input text
    Set oPPSlide = oPPPrsn.Slides(35)

    '------------------- CHINA -------------------
    Set oPPShape = oPPSlide.Shapes("FO-ChinaEur1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K110").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-ChinaEur2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L110").Value

    Set oPPShape = oPPSlide.Shapes("FO-ChinaEur3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M110").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-ChinaAus1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K111").Value

    Set oPPShape = oPPSlide.Shapes("FO-ChinaAus2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L111").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-ChinaAus3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M111").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-ChinaLatam1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K112").Value

    Set oPPShape = oPPSlide.Shapes("FO-ChinaLatam2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L112").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-ChinaLatam3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M112").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-ChinaAfrica1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K113").Value

    Set oPPShape = oPPSlide.Shapes("FO-ChinaAfrica2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L113").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-ChinaAfrica3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M113").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-MEChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K119").Value

    Set oPPShape = oPPSlide.Shapes("FO-MEChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L119").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-MEChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M119").Value
    
    '------------------- SINGAPORE -------------------
    Set oPPShape = oPPSlide.Shapes("FO-SGChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K114").Value

    Set oPPShape = oPPSlide.Shapes("FO-SGChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L114").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-SGChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M114").Value

    Set oPPShape = oPPSlide.Shapes("FO-EurSG1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K126").Value

    Set oPPShape = oPPSlide.Shapes("FO-EurSG2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L126").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-EurSG3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M126").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-LatamSG1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K127").Value

    Set oPPShape = oPPSlide.Shapes("FO-LatamSG2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L127").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-LatamSG3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M127").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-SGMsia1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K133").Value

    Set oPPShape = oPPSlide.Shapes("FO-SGMsia2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L133").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-SGMsia3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M133").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-SGKor1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K145").Value

    Set oPPShape = oPPSlide.Shapes("FO-SGKor2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L145").Value
    
    Set oPPShape = oPPSlide.Shapes("FO-SGKor3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M145").Value

    ' ----- NAPHTHA TRADE
    ' Slide with shape to input text
    Set oPPSlide = oPPPrsn.Slides(19)

    '------------------- MIDDLE EAST -------------------
    Set oPPShape = oPPSlide.Shapes("Nap-MEKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K153").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-MEKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L153").Value

    Set oPPShape = oPPSlide.Shapes("Nap-MEKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M153").Value

    Set oPPShape = oPPSlide.Shapes("Nap-MEChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K154").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-MEChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L154").Value

    Set oPPShape = oPPSlide.Shapes("Nap-MEChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M154").Value

    Set oPPShape = oPPSlide.Shapes("Nap-MEJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K155").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-MEJapan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L155").Value

    Set oPPShape = oPPSlide.Shapes("Nap-MEJapan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M155").Value

    Set oPPShape = oPPSlide.Shapes("Nap-METaiwan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K156").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-METaiwan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L156").Value

    Set oPPShape = oPPSlide.Shapes("Nap-METaiwan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M156").Value

    Set oPPShape = oPPSlide.Shapes("Nap-MESG1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K157").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-MESG2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L157").Value

    Set oPPShape = oPPSlide.Shapes("Nap-MESG3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M157").Value

    '------------------- INDIA -------------------
    Set oPPShape = oPPSlide.Shapes("Nap-IndiaKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K159").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-IndiaKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L159").Value

    Set oPPShape = oPPSlide.Shapes("Nap-IndiaKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M159").Value

    Set oPPShape = oPPSlide.Shapes("Nap-IndiaJap1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K160").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-IndiaJap2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L160").Value

    Set oPPShape = oPPSlide.Shapes("Nap-IndiaJap3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M160").Value

    Set oPPShape = oPPSlide.Shapes("Nap-IndiaTW1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K161").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-IndiaTW2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L161").Value

    Set oPPShape = oPPSlide.Shapes("Nap-IndiaTW3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M161").Value

    Set oPPShape = oPPSlide.Shapes("Nap-IndiaSG1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K162").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-IndiaSG2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L162").Value

    Set oPPShape = oPPSlide.Shapes("Nap-IndiaSG3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M162").Value

    '------------------- WEST OF SUEZ -------------------
    Set oPPShape = oPPSlide.Shapes("Nap-WestTW1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K164").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-WestTW2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L164").Value

    Set oPPShape = oPPSlide.Shapes("Nap-WestTW3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M164").Value

    Set oPPShape = oPPSlide.Shapes("Nap-WestChina1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K165").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-WestChina2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L165").Value

    Set oPPShape = oPPSlide.Shapes("Nap-WestChina3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M165").Value

    Set oPPShape = oPPSlide.Shapes("Nap-WestKorea1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K166").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-WestKorea2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L166").Value

    Set oPPShape = oPPSlide.Shapes("Nap-WestKorea3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M166").Value

    Set oPPShape = oPPSlide.Shapes("Nap-WestJapan1")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("K167").Value
    
    Set oPPShape = oPPSlide.Shapes("Nap-WestJapan2")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("L167").Value

    Set oPPShape = oPPSlide.Shapes("Nap-WestJapan3")
    oPPShape.TextFrame.TextRange.Text = _
    ThisWorkbook.Sheets("Sheet1").Range("M167").Value

End Sub




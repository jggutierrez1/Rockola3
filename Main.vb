Imports System.IO
Imports System.Management
Imports System.Management.Instrumentation
Imports System.Drawing
Imports Devart.Data.SQLite

Public Class FMain

    Private oGrpGene_With As Integer = 590
    Private oGrpGene_Height As Integer = 535
    Private oGrpGene_x As Integer = 379
    Private oGrpGene_y As Integer = 116

    Private oGrpDisc_With As Integer = 622
    Private oGrpDisc_Height As Integer = 601
    Private oGrpDisc_x As Integer = 380
    Private oGrpDisc_y As Integer = 127

    Private oGrpCanc_With As Integer = 590
    Private oGrpCanc_Height As Integer = 535
    Private oGrpCanc_x As Integer = 379
    Private oGrpCanc_y As Integer = 116
    Private bLoad_Array_Objects As Boolean = False


    Dim iShow As Integer
    Dim DriverOpened As Boolean
    Dim piCnt_Canc As Integer
    Dim bFlagChek1 As Boolean
    Dim bFlagChek2 As Boolean
    Dim bFlagChek3 As Boolean
    Dim gbServ_Mode As Boolean
    Dim bfTm As Boolean
    Dim iPosPagD As Integer
    Dim pTimeResto As Long
    Dim igLen2 As Integer
    Dim Transparent As Color = Nothing
    Dim iValue As Integer

    Dim Disc_Num As New Collection
    Dim Disc_Img As New Collection
    Dim Disc_Video As New Collection
    Dim Disk_New As New Collection
    Dim oDisc_LabelA As New Collection
    Dim oDisc_LabelB As New Collection
    Dim Canc_Nom As New Collection
    Dim Canc_Video As New Collection
    Dim Gene_Lst As New Collection

    Private Sub Load_Array_Objects()

        Gene_Lst.Clear()
        Gene_Lst.Add(Me.oLGenero1)
        Gene_Lst.Add(Me.oLGenero2)
        Gene_Lst.Add(Me.oLGenero3)
        Gene_Lst.Add(Me.oLGenero4)
        Gene_Lst.Add(Me.oLGenero5)
        Gene_Lst.Add(Me.oLGenero6)
        Gene_Lst.Add(Me.oLGenero7)
        Gene_Lst.Add(Me.oLGenero8)
        Gene_Lst.Add(Me.oLGenero9)
        Gene_Lst.Add(Me.oLGenero10)
        Gene_Lst.Add(Me.oLGenero11)
        Gene_Lst.Add(Me.oLGenero12)
        Gene_Lst.Add(Me.oLGenero13)

        Disc_Num.Clear()
        Disc_Num.Add(Me.Disc_Num1)
        Disc_Num.Add(Me.Disc_Num2)
        Disc_Num.Add(Me.Disc_Num3)
        Disc_Num.Add(Me.Disc_Num4)
        Disc_Num.Add(Me.Disc_Num5)
        Disc_Num.Add(Me.Disc_Num6)
        Disc_Num.Add(Me.Disc_Num7)
        Disc_Num.Add(Me.Disc_Num8)
        Disc_Num.Add(Me.Disc_Num9)
        Disc_Num.Add(Me.Disc_Num10)
        Disc_Num.Add(Me.Disc_Num11)
        Disc_Num.Add(Me.Disc_Num12)

        Disk_New.Clear()
        Disk_New.Add(Me.Disk_New1)
        Disk_New.Add(Me.Disk_New2)
        Disk_New.Add(Me.Disk_New3)
        Disk_New.Add(Me.Disk_New4)
        Disk_New.Add(Me.Disk_New5)
        Disk_New.Add(Me.Disk_New6)
        Disk_New.Add(Me.Disk_New7)
        Disk_New.Add(Me.Disk_New8)
        Disk_New.Add(Me.Disk_New9)
        Disk_New.Add(Me.Disk_New10)
        Disk_New.Add(Me.Disk_New11)
        Disk_New.Add(Me.Disk_New12)

        Disc_Video.Clear()
        Disc_Video.Add(Me.Disc_Video1)
        Disc_Video.Add(Me.Disc_Video2)
        Disc_Video.Add(Me.Disc_Video3)
        Disc_Video.Add(Me.Disc_Video4)
        Disc_Video.Add(Me.Disc_Video5)
        Disc_Video.Add(Me.Disc_Video6)
        Disc_Video.Add(Me.Disc_Video7)
        Disc_Video.Add(Me.Disc_Video8)
        Disc_Video.Add(Me.Disc_Video9)
        Disc_Video.Add(Me.Disc_Video10)
        Disc_Video.Add(Me.Disc_Video11)
        Disc_Video.Add(Me.Disc_Video12)

        Disc_Img.Clear()
        Disc_Img.Add(Me.Disc_Img1)
        Disc_Img.Add(Me.Disc_Img2)
        Disc_Img.Add(Me.Disc_Img3)
        Disc_Img.Add(Me.Disc_Img4)
        Disc_Img.Add(Me.Disc_Img5)
        Disc_Img.Add(Me.Disc_Img6)
        Disc_Img.Add(Me.Disc_Img7)
        Disc_Img.Add(Me.Disc_Img8)
        Disc_Img.Add(Me.Disc_Img9)
        Disc_Img.Add(Me.Disc_Img10)
        Disc_Img.Add(Me.Disc_Img11)
        Disc_Img.Add(Me.Disc_Img12)

        oDisc_LabelA.Clear()
        oDisc_LabelA.Add(Me.oDisc_LabelA1)
        oDisc_LabelA.Add(Me.oDisc_LabelA2)
        oDisc_LabelA.Add(Me.oDisc_LabelA3)
        oDisc_LabelA.Add(Me.oDisc_LabelA4)
        oDisc_LabelA.Add(Me.oDisc_LabelA5)
        oDisc_LabelA.Add(Me.oDisc_LabelA6)
        oDisc_LabelA.Add(Me.oDisc_LabelA7)
        oDisc_LabelA.Add(Me.oDisc_LabelA8)
        oDisc_LabelA.Add(Me.oDisc_LabelA9)
        oDisc_LabelA.Add(Me.oDisc_LabelA10)
        oDisc_LabelA.Add(Me.oDisc_LabelA11)
        oDisc_LabelA.Add(Me.oDisc_LabelA12)

        oDisc_LabelB.Clear()
        oDisc_LabelB.Add(Me.oDisc_LabelB1)
        oDisc_LabelB.Add(Me.oDisc_LabelB2)
        oDisc_LabelB.Add(Me.oDisc_LabelB3)
        oDisc_LabelB.Add(Me.oDisc_LabelB4)
        oDisc_LabelB.Add(Me.oDisc_LabelB5)
        oDisc_LabelB.Add(Me.oDisc_LabelB6)
        oDisc_LabelB.Add(Me.oDisc_LabelB7)
        oDisc_LabelB.Add(Me.oDisc_LabelB8)
        oDisc_LabelB.Add(Me.oDisc_LabelB9)
        oDisc_LabelB.Add(Me.oDisc_LabelB10)
        oDisc_LabelB.Add(Me.oDisc_LabelB11)
        oDisc_LabelB.Add(Me.oDisc_LabelB12)

        Canc_Nom.Clear()
        Canc_Nom.Add(Me.oLCanc1)
        Canc_Nom.Add(Me.oLCanc2)
        Canc_Nom.Add(Me.oLCanc3)
        Canc_Nom.Add(Me.oLCanc4)
        Canc_Nom.Add(Me.oLCanc5)
        Canc_Nom.Add(Me.oLCanc6)
        Canc_Nom.Add(Me.oLCanc7)
        Canc_Nom.Add(Me.oLCanc8)
        Canc_Nom.Add(Me.oLCanc9)
        Canc_Nom.Add(Me.oLCanc10)
        Canc_Nom.Add(Me.oLCanc11)
        Canc_Nom.Add(Me.oLCanc12)
        Canc_Nom.Add(Me.oLCanc13)
        Canc_Nom.Add(Me.oLCanc14)
        Canc_Nom.Add(Me.oLCanc15)

        Canc_Video.Clear()
        Canc_Video.Add(Me.oICancVideo1)
        Canc_Video.Add(Me.oICancVideo2)
        Canc_Video.Add(Me.oICancVideo3)
        Canc_Video.Add(Me.oICancVideo4)
        Canc_Video.Add(Me.oICancVideo5)
        Canc_Video.Add(Me.oICancVideo6)
        Canc_Video.Add(Me.oICancVideo7)
        Canc_Video.Add(Me.oICancVideo8)
        Canc_Video.Add(Me.oICancVideo9)
        Canc_Video.Add(Me.oICancVideo10)
        Canc_Video.Add(Me.oICancVideo11)
        Canc_Video.Add(Me.oICancVideo12)
        Canc_Video.Add(Me.oICancVideo13)
        Canc_Video.Add(Me.oICancVideo14)
        Canc_Video.Add(Me.oICancVideo15)

        bLoad_Array_Objects = True
    End Sub

    Private Sub Gen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
            oLGenero1.Click, oLGenero2.Click, oLGenero3.Click, oLGenero4.Click, oLGenero5.Click,
        oLGenero6.Click, oLGenero7.Click, oLGenero8.Click, oLGenero9.Click, oLGenero10.Click,
        oLGenero11.Click, oLGenero12.Click, oLGenero13.Click

        If Not (sender.tag = 0) Then
            igGen_Sel = sender.tag
            igAct_PgD = 1
            Call Cargar_Pag_Dis(sender.tag, 1)
        End If
    End Sub

    Private Sub Disc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        Disc_Img1.Click, Disc_Img2.Click, Disc_Img3.Click, Disc_Img4.Click, Disc_Img5.Click,
        Disc_Img6.Click, Disc_Img7.Click, Disc_Img8.Click, Disc_Img9.Click, Disc_Img10.Click,
        Disc_Img11.Click, Disc_Img12.Click

        If Not (sender.tag = 0) Then
            igDis_Sel = sender.tag
            igAct_PgC = 1
            Call Cargar_Pag_Canc(sender.tag, 1)
        End If
    End Sub

    Private Sub Canc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
        oLCanc1.Click, oLCanc2.Click, oLCanc3.Click, oLCanc4.Click, oLCanc5.Click,
        oLCanc6.Click, oLCanc7.Click, oLCanc8.Click, oLCanc9.Click, oLCanc10.Click,
        oLCanc11.Click, oLCanc12.Click, oLCanc13.Click

        If Not (sender.tag = 0) Then
            igCan_Sel = sender.tag

            igAct_PgG = 1
            Cargar_Pag_Gen(igAct_PgG)
        End If
    End Sub

    Sub Desactiva_Genero(ByRef pFlag As Boolean)
        Me.oGrpBox_Gen.BackColor = Color.Transparent
        Me.oGrpBox_Gen.Refresh()
        Me.oGrpBox_Gen.Visible = Not pFlag

    End Sub

    Sub Desactiva_Disco(ByRef pFlag As Boolean)
        Me.oGrpBox_Dis.BackColor = Color.Transparent
        Me.oGrpBox_Dis.Refresh()
        Me.oGrpBox_Dis.Visible = Not pFlag
    End Sub

    Sub Desactiva_Cancion(ByRef pFlag As Boolean)
        Me.oGrpBox_Can.BackColor = Color.Transparent
        Me.oGrpBox_Can.Refresh()
        Me.oGrpBox_Can.Visible = Not pFlag
    End Sub

    Public Sub Limpia_Gen()
        Dim i As Integer
        For i = 1 To Me.Gene_Lst.Count
            Me.Gene_Lst(i).Text = ""
            Me.Gene_Lst(i).tag = 0
            Me.Gene_Lst(i).BorderStyle = BorderStyle.None
        Next i
    End Sub

    Public Sub Limpia_Dis()
        Dim i As Integer
        Dim bVisible As Boolean

        For i = 1 To Disc_Num.Count
            'NUMERO DE DISCO
            Disc_Num(i).text = ""
            Disc_Num(i).Visible = False
            Disc_Num(i).cursor = Cursors.Default
            Disc_Num(i).BorderStyle = BorderStyle.None
            Disc_Num(i).BackColor = Color.Transparent
            Disc_Num(i).TAG = 0

            'LEYENDAS DE ABAJO A
            oDisc_LabelA(i).text = ""
            oDisc_LabelA(i).cursor = Cursors.Default
            oDisc_LabelA(i).BorderStyle = BorderStyle.None
            oDisc_LabelA(i).Visible = False

            'LEYENDAS DE ABAJO B
            oDisc_LabelB(i).text = ""
            oDisc_LabelB(i).Visible = False
            oDisc_LabelB(i).cursor = Cursors.Default
            oDisc_LabelB(i).BorderStyle = BorderStyle.None

            'IMAGEN DEL DISCO
            Disc_Img(i).image = Nothing
            Disc_Img(i).imagelocation = ""
            Disc_Img(i).TAG = 0

            'fMain.oImg_Disc(i).Cursors = Cursors.Default

            'ETIQUETA DE DISCO NUEVO
            Disk_New(i).Visible = False

            'ETIQUETA DE DISCO CON VIDEO
            Disc_Video(i).Visible = False

        Next i
    End Sub

    Public Sub Limpia_Can()
        Dim i As Integer
        For i = 1 To Canc_Nom.Count
            Canc_Nom(i).text = ""
            Canc_Nom(i).tag = 0
            'fMain.oLCanc(i).Cursors = Cursors.Default
            Canc_Video(i).Visible = False
        Next i
    End Sub

    Private Sub Cargar_Gen()
        Call Me.Cargar_Pag_Gen(1)
    End Sub

    Private Sub Cargar_Pag_Gen(ByVal pNum_Pag As Integer)
        If bLoad_Array_Objects = False Then
            Me.Load_Array_Objects()
        End If

        Me.Desactiva_Genero(True)
        Me.Desactiva_Disco(True)
        Me.Desactiva_Cancion(True)

        '------------------------LIMPIA DATOS EN PANTALLA---------------------------------
        REM Me.Limpia_Gen()
        '----------------------------------------------------------------------------------

        igAct_PgD = 1
        igAct_PgC = 1
        Call Max_Pag_Gen()

        Dim iNum_reg As Integer = 0
        '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------
        Dim sql As String = "SELECT id_gen, id_ord, descri FROM file01 WHERE page=" & pNum_Pag & " ORDER BY id_ord"
        Dim dar As SQLiteDataReader

        If Not (oConnSqLite.State = ConnectionState.Open) Then
            oConnSqLite.Open()
        End If

        oCmmSqLite = New SQLiteCommand(sql, oConnSqLite)
        dar = oCmmSqLite.ExecuteReader()
        '----------------------------------------------------------------------------------

        Me.oLTitulo.Text = "SELECCIONE EL GENERO."

        '------------------------MOSTRAR DATOS EN PANTALLA---------------------------------
        Dim i As Integer = 0
        If dar.HasRows = True Then
            iNum_reg = 1
            Do While dar.Read()
                Me.Gene_Lst(iNum_reg).tag = Val(dar.Item("id_gen").ToString())
                Me.Gene_Lst(iNum_reg).text = dar.Item("id_ord").ToString & "-" & dar.Item("descri").ToString
                Me.Gene_Lst(iNum_reg).cursor = Cursors.Hand
                Me.Gene_Lst(iNum_reg).visible = True
                Me.Gene_Lst(iNum_reg).refresh
                iNum_reg = iNum_reg + 1
            Loop
        Else
            iNum_reg = 1
        End If
        '----------------------------------------------------------------------------------
        For i = iNum_reg To Me.Gene_Lst.Count
            Me.Gene_Lst(i).Text = ""
            Me.Gene_Lst(i).tag = 0
            Me.Gene_Lst(i).BorderStyle = BorderStyle.None
            Me.Gene_Lst(i).cursor = Cursors.Default
            Me.Gene_Lst(i).visible = True

            Me.Gene_Lst(i).refresh
        Next i
        '----------------------------------------------------------------------------------
        dar.Close()
        dar = Nothing
        Me.Desactiva_Genero(False)
    End Sub

    Private Sub Cargar_Pag_Dis(ByVal pCod_Gen As String, ByVal pNum_Pag As Integer)
        Me.Desactiva_Genero(True)
        Me.Desactiva_Disco(True)
        Me.Desactiva_Cancion(True)

        '------------------------LIMPIA DATOS EN PANTALLA---------------------------------
        REM Me.Limpia_Dis()
        '----------------------------------------------------------------------------------

        igAct_PgC = 1
        Call Max_Pag_Dis(pCod_Gen)

        Dim iNum_reg As Integer = 0
        '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------
        Dim sql As String = "SELECT * FROM file02 WHERE id_gen=" & pCod_Gen & " AND page=" & pNum_Pag & " ORDER BY id_gen, id_ord"
        Dim dar As SQLiteDataReader

        If Not (oConnSqLite.State = ConnectionState.Open) Then
            oConnSqLite.Open()
        End If

        oCmmSqLite = New SQLiteCommand(sql, oConnSqLite)
        dar = oCmmSqLite.ExecuteReader()
        '----------------------------------------------------------------------------------

        Me.oLTitulo.Text = "[SELECCIONE EL DISCO.]"

        '------------------------MOSTRAR DATOS EN PANTALLA---------------------------------
        Dim sFile As String = ""
        Dim sLabel As String = ""
        Dim i As Integer = 0
        Dim cPahFile As String = ""
        Dim cFldFoto As String = ""

        If ((dar.HasRows = True) And (dar.EndOfData = False)) Then
            iNum_reg = 1
            Do While dar.Read()
                cFldFoto = Strings.Trim(dar.Item("FL_IMG").ToString())
                sFile = Trim(System.IO.Path.GetFileName(cFldFoto))
                cPahFile = sParam(4) + "\" + Trim(System.IO.Path.GetFileName(cFldFoto))
                sLabel = Strings.Trim(dar.Item("ID_ORD").ToString)
                If Strings.Right(cPahFile, 1) = "\" Then
                    cPahFile = ""
                End If

                Me.Disc_Num(iNum_reg).TAG = Val(dar.Item("id_dis").ToString())
                Me.Disc_Num(iNum_reg).text = sLabel
                Me.Disc_Num(iNum_reg).cursor = Cursors.Hand
                Me.Disc_Num(iNum_reg).Visible = True

                Me.Disc_Img(iNum_reg).TAG = Val(dar.Item("id_dis").ToString())
                Me.Disc_Img(iNum_reg).cursor = Cursors.Hand

                If (Strings.Trim(sFile) = "") Then
                    Me.Disc_Img(iNum_reg).ImageLocation = ""
                    Me.Disc_Img(iNum_reg).Image = My.Resources.Resources.sin_images
                Else
                    'If IO.File.Exists(cPahFile) = True Then
                    Me.Disc_Img(iNum_reg).Image = Nothing
                    Me.Disc_Img(iNum_reg).ImageLocation = cPahFile
                    'Else
                    'Me.Disc_Img(iNum_reg).ImageLocation = ""
                    'Me.Disc_Img(iNum_reg).Image = My.Resources.Resources.no_images
                    'End If
                End If

                Me.Disc_Img(iNum_reg).Visible = True

                If Val(dar.Item("C_VIDEO").ToString) = 1 Then
                    If bgVideoLabel = True Then
                        Me.Disc_Video(iNum_reg).Visible = True
                    End If
                Else
                    Me.Disc_Video(iNum_reg).Visible = False
                End If

                If Val(dar.Item("FL_NEW").ToString) = 1 Then
                    Me.Disk_New(iNum_reg).Visible = True
                Else
                    Me.Disk_New(iNum_reg).Visible = False
                End If

                If bgDiscLabel = True Then
                    Me.oDisc_LabelA(iNum_reg).Visible = True
                    Me.oDisc_LabelA(iNum_reg).text = UCase(Strings.Trim(PADL(dar.Item("NOM_DIS").ToString, 19, " ")))
                    Me.oDisc_LabelB(iNum_reg).Visible = True
                    Me.oDisc_LabelB(iNum_reg).text = UCase(Strings.Trim(PADL(dar.Item("NOM_ART").ToString, 19, " ")))
                Else
                    Me.oDisc_LabelA(iNum_reg).text = ""
                    Me.oDisc_LabelA(iNum_reg).Visible = False
                    Me.oDisc_LabelB(iNum_reg).text = ""
                    Me.oDisc_LabelB(iNum_reg).Visible = False
                End If

                Me.Disc_Num(iNum_reg).refresh
                Me.oDisc_LabelA(iNum_reg).refresh
                Me.oDisc_LabelB(iNum_reg).refresh
                Me.Disc_Img(iNum_reg).refresh
                Me.Disk_New(iNum_reg).refresh
                Me.Disc_Video(iNum_reg).refresh
                iNum_reg = iNum_reg + 1
            Loop
        Else
            iNum_reg = 1
        End If
        '----------------------------------------------------------------------------------
        For i = iNum_reg To Me.Disc_Img.Count
            'NUMERO DE DISCO
            Me.Disc_Num(i).text = ""
            Me.Disc_Num(i).Visible = False
            Me.Disc_Num(i).cursor = Cursors.Default
            Me.Disc_Num(i).BorderStyle = BorderStyle.None
            Me.Disc_Num(i).BackColor = Color.Transparent

            'LEYENDAS DE ABAJO A
            Me.oDisc_LabelA(i).text = ""
            Me.oDisc_LabelA(i).BorderStyle = BorderStyle.None
            Me.oDisc_LabelA(i).Visible = False

            'LEYENDAS DE ABAJO B
            Me.oDisc_LabelB(i).text = ""
            Me.oDisc_LabelB(i).Visible = False
            Me.oDisc_LabelB(i).BorderStyle = BorderStyle.None

            'IMAGEN DEL DISCO
            Me.Disc_Img(i).image = Nothing
            Me.Disc_Img(i).imagelocation = ""
            Me.Disc_Img(i).cursor = Cursors.Default
            Me.Disc_Img(i).Visible = False

            'fMain.oImg_Disc(i).Cursors = Cursors.Default

            'ETIQUETA DE DISCO NUEVO
            Me.Disk_New(i).Visible = False

            'ETIQUETA DE DISCO CON VIDEO
            Me.Disc_Video(i).Visible = False

            Me.Disc_Num(i).refresh
            Me.oDisc_LabelA(i).refresh
            Me.oDisc_LabelB(i).refresh
            Me.Disc_Img(i).refresh
            Me.Disk_New(i).refresh
            Me.Disc_Video(i).refresh
        Next i
        '----------------------------------------------------------------------------------
        dar.Close()
        dar = Nothing
        Me.Desactiva_Disco(False)
    End Sub

    Private Sub Cargar_Pag_Canc(ByVal pCod_Dis As String, ByVal pNum_Pag As Integer)
        Me.Desactiva_Genero(True)
        Me.Desactiva_Disco(True)
        Me.Desactiva_Cancion(True)

        '------------------------LIMPIA DATOS EN PANTALLA---------------------------------
        REM Me.Limpia_Can()
        '----------------------------------------------------------------------------------

        Call Max_Pag_Canc(pCod_Dis)

        Dim iNum_reg As Integer = 0
        '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------
        Dim sql As String = "SELECT * FROM file03 WHERE id_dis=" & pCod_Dis & " AND page=" & pNum_Pag & " ORDER BY id_dis, id_ord"
        Dim dar As SQLiteDataReader

        If Not (oConnSqLite.State = ConnectionState.Open) Then
            oConnSqLite.Open()
        End If

        oCmmSqLite = New SQLiteCommand(sql, oConnSqLite)
        dar = oCmmSqLite.ExecuteReader()
        '----------------------------------------------------------------------------------

        Me.oLTitulo.Text = "[SELECCIONE EL TEMA.]"

        '------------------------MOSTRAR DATOS EN PANTALLA---------------------------------
        Dim sOrder As String = ""
        Dim sLabel As String = ""
        Dim sNamFl As String = ""
        Dim sExtFl As String = ""
        Dim sFile As String = ""
        Dim i As Integer = 0

        If dar.HasRows = True Then
            iNum_reg = 1
            Do While dar.Read()

                sOrder = dar.Item("ID_ORD").ToString
                sLabel = dar.Item("DE_CAN").ToString
                sNamFl = System.IO.Path.GetFileName(Strings.Trim(dar.Item("fl_mp3").ToString))
                sExtFl = System.IO.Path.GetExtension(Strings.Trim(dar.Item("fl_mp3").ToString))

                Me.Canc_Video(iNum_reg).Visible = True
                If Not (sExtFl <> "MP3") Or (sExtFl <> ".MP3") Then
                    Me.Canc_Video(iNum_reg).Image = My.Resources.Resources.icn_video_pk
                Else
                    Me.Canc_Video(iNum_reg).Image = My.Resources.Resources.themes
                End If
                Me.Canc_Video(iNum_reg).Visible = True

                If Not sLabel = "" Then
                    sLabel = sOrder & " - " & Proper(sLabel)
                End If
                sFile = sgDir_Mp3 & sNamFl

                If IO.File.Exists(sFile) = True Then
                    Me.Canc_Nom(iNum_reg).Font = New Font(Me.oLCanc1.Font.Name, Me.oLCanc1.Font.Size, FontStyle.Bold Or Not FontStyle.Strikeout)
                Else
                    Me.Canc_Nom(iNum_reg).Font = New Font(Me.oLCanc1.Font.Name, Me.oLCanc1.Font.Size, FontStyle.Bold Or FontStyle.Strikeout)
                End If

                Me.Canc_Nom(iNum_reg).TAG = dar.Item("id_can").ToString
                Me.Canc_Nom(iNum_reg).TEXT = sLabel
                Me.Canc_Nom(iNum_reg).cursor = Cursors.Hand
                Me.Canc_Nom(iNum_reg).Visible = True

                Me.Canc_Nom(iNum_reg).refresh
                Me.Canc_Video(iNum_reg).refresh
                iNum_reg = iNum_reg + 1
            Loop
        Else
            iNum_reg = 1
        End If
        '----------------------------------------------------------------------------------
        For i = iNum_reg To Me.Canc_Nom.Count
            Me.Canc_Nom(i).text = ""
            Me.Canc_Nom(i).tag = 0
            Me.Canc_Nom(i).cursor = Cursors.Default
            Me.Canc_Nom(i).Visible = False

            Me.Canc_Video(i).Image = Nothing
            Me.Canc_Video(i).Visible = False

            Me.Canc_Nom(i).refresh
            Me.Canc_Video(i).refresh
        Next i
        '----------------------------------------------------------------------------------

        Me.Desactiva_Cancion(False)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        End
    End Sub

    Private Sub FMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If bLoad_Array_Objects = False Then
            Me.Load_Array_Objects()
        End If
        Dim MiValor As String = ""
        sgParms = Environment.GetCommandLineArgs()

        If sgParms.Count() > -1 Then
            If sgParms(0) = "ACTIVATE" Then
                If UBound(sgParms) = 0 Then
                    MiValor = InputBox("Inserte el código de seguridad", "Recreativos Veraguenses", "", 100, 100)
                    If Val(MiValor) <> 2527 Then
                        End
                    Else
                        REM Act_Form1.Show vbModal
                        End
                    End If
                ElseIf UBound(sgParms) = 1 Then
                    If sgParms(1) = "2527" Then
                        REM Act_Form1.Show vbModal
                        End
                    End If
                End If
            End If
        End If

        Me.oGrpBox_Can.Visible = False
        Me.oGrpBox_Dis.Visible = False
        Me.oGrpBox_Gen.Visible = False

        bfTm = False
        igCnt_CR = 0 : igCnt_CRP = 0 : igCnt_CRG = 0
        igDelay_Ret_Gen = 0
        igStartPlayMusic = 0
        gbServ_Mode = False
        igFlg_TCR = 0
        igNoDuplicT = 0
        igInd_Kar = 0
        igInd_Pub = 0
        igDelay_Del_Dig_Can = 0
        igLen2 = 0 : igNo_RgAt = 0
        igAct_PgG = 1 : igTot_PgG = 8 : igTot_PgC = 8
        igAct_PgD = 1 : igTot_PgD = 8 : igTot_PgC = 0
        igMax_Gen = 13 - 2 : igMax_Dis = 12 : igMax_Can = 13 'valores fijos su valor no pueded ser superior
        igInd_Bon = 0
        igNext_Bonus = 0
        piCnt_Canc = 0
        igShowPass = 0
        bFlagFoc = True
        bgBlinkPag = False
        bgPopular = False
        bgVIP = False
        bgIs_Video = False
        bgIs_Publi = False
        bgWMP_Busy = False
        igCont_Sin = 0
        Me.oLVersion.Text = "Ver." & Strings.Trim(Application.ProductVersion)

        If My.Computer.Keyboard.NumLock = False Then
            My.Computer.Keyboard.SendKeys("{NUMLOCK}")
        End If

        Me.Colocar_Frames()
        Call Init_Ini()
        REM Call Save_Defa_Path()
        Call Get_System_Path(sParam)

        '---------------------Carga Variables de rutas----------------------------
        Dim iTmp As String
        Dim sTmp As String

        sgDir_odb = sParam(1)
        sgDir_Tmp = sParam(2)
        sgDir_Fls = sParam(3)
        sgDir_Img = sParam(4)
        sgDir_Mp3 = sParam(5)
        sgDir_Pub1 = sParam(6)
        sgFec_iAc = sParam(7)
        sgFec_Fac = sParam(8)
        sgSer_Mac = sParam(9)
        sgNom_Loc = Strings.Trim(sParam(10))
        sgFle_Fon = sParam(13)
        sgSer_CPU = sParam(14)
        sTmp = oIni.GetKeyValue("CREDITS", "CEDIT_ACAC", "0")
        igCnt_CRG = Val(sTmp)
        REM igCnt_CRG = Val(UnScramble(sTmp).ToString())
        igFlg_TCR = Val(oIni.GetKeyValue("CREDITS", "FLG_TESTCR", "0"))
        igDelay_Return_Gen = Int(Val(sParam(15)))
        igDelay_Return_Dis = Int(Val(sParam(16)))
        igDelay_Bonus_Vid = Int(Val(sParam(17)))
        igNext_Bonus = (Today.Hour * 60) + (Today.Minute() + igDelay_Bonus_Vid)
        igLim_Cred = Int(Val(sParam(20)))
        igKeep_Cred = Int(Val(sParam(21)))
        If (igKeep_Cred = 0) Then
            If igFlg_SavedCR = 1 Then
                igCnt_CR = Val(oIni.GetKeyValue("CREDITS", "ACU_SAVECR", "#####0"))
            Else
                igCnt_CR = 0
            End If
        Else
            igCnt_CR = 6
        End If
        If igFlg_TCR = 1 Then
            Me.olMetros2.Visible = True
            Me.olMetros2.Text = PADL(0, 6, "0")
            Me.olLabMetros2.Visible = True
        Else
            Me.olMetros2.Visible = False
            Me.olLabMetros2.Visible = False
        End If
        igMixe_Popu = Int(Val(sParam(22)))
        sgKb_Crd1 = sParam(23)
        sgKb_Crd2 = sParam(25)
        sgKb_Del = sParam(27)
        sgKb_Ret = sParam(29)
        sgKb_ResM = sParam(31)
        sgKb_ResA = sParam(33)
        sgKb_Pop = sParam(35)
        sgKb_VID = sParam(36)
        sgKb_VIP = sParam(37)
        sgKb_BonC = Int(Val(sParam(38)))
        sgKb_UP = sParam(39)
        sgKb_Vef = sParam(40)
        sgKb_DN = sParam(41)
        sgKb_SwP = sParam(51)
        sgWin_Key = sParam(43)
        sgCr_AKey = sParam(52)
        sgIdx_Prm = sParam(53)
        bgSw = True
        bgVideoLabel = IIf(Val(sParam(44)) = 1, True, False)
        bgDiscLabel = IIf(Val(sParam(45)) = 1, True, False)
        bgKeep_On_Top = IIf(sParam(46) = 0, False, True)
        igScr_Alone = Int(Val(sParam(47)))
        igNoDuplicT = Int(Val(sParam(48)))
        bgSw_Pub = IIf(Val(Int(sParam(49))) = 1, True, False)

        If bgSw_Pub = False Then
            Me.oInd_VideoSW.Image = My.Resources.Resources.grnled
        Else
            Me.oInd_VideoSW.Image = My.Resources.Resources.redled
        End If
        Me.oInd_VideoSW.Visible = True

        sgDir_Pub2 = sParam(50)
        sgWin_Key = sParam(43)
        igLeftDisk = 700
        If UBound(sgParms) > -1 Then
            Select Case sgParms(0)
                Case Is = "CONVERSION"
                    REM Call Set_Open_Dbf
                    REM sPath = IIf(igInd_Kar = 0, sgDir_Fls, sgDir_Fls2)
                    REM Call ogVFP9.EXPORT__TABFILES(sPath, sPath, sgDir_Tmp)
                    Call MsgBox("La conversión se realizó:..")
                    End
                Case Is = "SERVICE"
                    If UBound(sgParms) = 0 Then
                        MiValor = InputBox("Inserte el código de seguridad", "Flamingo Magic Game", "", 100, 100)
                        If Val(MiValor) <> 2527 Then
                            End
                        Else
                            REM Call Go_Service
                            End
                        End If
                    Else
                        If sgParms(1) = 2527 Then
                            REM Call Go_Service
                            End
                        End If
                    End If
            End Select
        End If
        Call Check_Other()
        REM On Error GoTo Solve_error
        If sgNom_Loc <> UCase("SIN ASIGNACIÓN!") Then
            Me.olLocal.Text = "[" & sgNom_Loc & "] "
        Else
            Me.olLocal.Text = ""
        End If
        Me.olMensajeSis.Text = ""
        Me.olLocal.Text = Me.olLocal.Text & sgWin_Key
        Dim iReadFlg As Integer

        'iReadFlg = Int(Val(oIni.GetKeyValue("ROCKOLA", "RUNNINGSEC", "0")))
        'If iReadFlg = 1 Then
        '    Dim sVal1 As String
        '    Dim sVal2 As String
        '    If CheckForKL = False Then
        '        Call KTASK(TERMINATE, 0, 0, 0)
        '        End
        '    Else
        '        Call KTASK(READAUTH, READCODE1, READCODE2, READCODE3)
        '        oIni.SetKeyValue("SERIAL", "ID", Scramble("5C01001080000000000666413036").ToString)
        '    End If
        '    sVal1 = Scramble(Strings.Trim(ReadText()))
        '    sVal2 = Strings.Trim(oIni.GetKeyValue("SERIAL", "ID", ""))
        '    If (sVal1 <> sVal1) Then
        '        MsgBox("USB-KEY device no match.", vbCritical)
        '        End
        '    End If
        '    Me.olActivacion.Text = "Próxima activación: UNLIMITED WITH USB-KEY"
        'Else
        '    Dim sTmp1 As String
        '    Dim iValue As Integer
        '    Dim sValue As String
        '    Dim sMensage As String
        '    sMensage = "La copia del sistema no ha sido debidamente instalada o no ha sido activada"
        '    iValue = Int(Val(oIni.GetKeyValue("GENERAL", "NDISCVALID", "0")))
        '    If iValue = 0 Then
        '        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
        '        sTmp1 = allDrives(0).VolumeLabel
        '        sTmp1 = Strings.Left(sTmp1, 4) & "-" & Strings.Right(sTmp1, 4)
        '        If sgSer_Mac <> sTmp1 Then
        '            Call MsgBox(sMensage, vbCritical, "El sistema a sido movido de DISCO")
        '            REM Call Call_Tsr
        '            End
        '        End If
        '    End If
        '    iValue = Int(Val(oIni.GetKeyValue("GENERAL", "NMOTHVALID", "0")))
        '    If iValue = 0 Then
        '        sTmp1 = MBCPUNumber()
        '        If sgSer_CPU <> sTmp1 Then
        '            Call MsgBox(sMensage, vbCritical, "El sistema a sido movido de MÁQUINA")
        '            REM Call Call_Tsr
        '            End
        '        End If
        '    End If
        '    iValue = Int(Val(oIni.GetKeyValue("GENERAL", "NDATEVALID", "0")))
        '    If iValue = 0 Then
        '        If DateValue(sgFec_iAc) = DateValue(sgFec_Fac) Then
        '            Call MsgBox(sMensage, vbCritical, "La copia del sistemas debe ser activada")
        '            REM Call Call_Tsr
        '            End
        '        End If
        '        If Date.Now < DateValue(sgFec_iAc) Then
        '            Call MsgBox(sMensage, vbCritical, "La copia del sistema ha perdido vigencia")
        '            REM Call Call_Tsr
        '            End
        '        End If
        '        If Date.Now > DateValue(sgFec_Fac) Then
        '            Call MsgBox(sMensage, vbCritical, "La copia del sistema ha perdido vigencia")
        '            REM Call Call_Tsr
        '            End
        '        End If
        '    End If
        '    Me.olActivacion.Text = "Próxima activación: " & Format(sgFec_Fac, "dd/MM/yyyy")
        'End If
        iShow = Int(Val(oIni.GetKeyValue("GENERAL", "SHOW_MOUSE", "0")))
        Call ShowCursor(IIf(iShow = 1, True, False))

        '------------------------------------------------------------
        Dim sWinDir As String
        sWinDir = GetTheWindowsDirectory()

        Call SqlLite_Verify_Data()

        Err.Clear()
        Me.oLst_Temas_Video.Items.Clear()
        '"*.mpg|*.wmv|*.mp4|*.mpeg|*.mov|*.avi"
        If My.Computer.FileSystem.DirectoryExists(sgDir_Mp3) = True Then
            For Each foundFile As String In Directory.GetFiles(sgDir_Mp3, "*.mpg", SearchOption.TopDirectoryOnly)
                Me.oLst_Temas_Video.Items.Add(foundFile)
            Next
        End If

        Me.oGrpBox_Dis.BackColor = Transparent
        Me.oGrpBox_Gen.BackColor = Transparent
        Me.oGrpBox_Can.BackColor = Transparent

        Me.oTM_codigo2.Interval = igDelay_Return_Dis * 1000
        Me.oTM_codigo2.Enabled = False
        Me.Refresh_Creditos()
        '----------------------GENEROS------------------------------

        Me.Cargar_Gen()
        Me.Desactiva_Genero(False)
        Me.otCodigo.Focus()

        '---------------------PUBLICIDAD----------------------------
        Me.Conectar_DBPub()
        '-----------------------PROMOS------------------------------
        Me.Conectar_DBPromo()
        '-----------------------Otros-------------------------------
        'Me.otNot_Found_List.Text = ""


        Me.Cargar_Temas()
        If igScr_Alone = 0 Then
            Video_Form.Show()
        End If
        If bgKeep_On_Top = True Then
            ' Call AlwaysOnTop(Main_Form, True)
        End If
        If igStartPlayMusic = 1 Then
            My.Computer.Keyboard.SendKeys(sgKb_ResA)
            Me.Salvar_Temas()
        End If
        Exit Sub

Solve_error:
        Call ControlError()
        Resume Next
    End Sub

    Private Sub Check_Other()
        If Not sgFle_Fon = "" Then
            If IO.File.Exists(sgFle_Fon) Then
                Me.BackgroundImage = System.Drawing.Image.FromFile(sgFle_Fon)
                'aqui
            End If
        End If
        On Error Resume Next
        My.Computer.FileSystem.CreateDirectory(sgDir_Tmp)
        On Error Resume Next
        My.Computer.FileSystem.CreateDirectory(sgDir_odb)
        '---------------------------------------Link_Tab.dsn---------------------------------------
        Dim oIni2 As New IniFile

        oIni2.Load(sgDir_odb & "\Link_Tab.dsn")
        If Not IO.File.Exists(sgDir_odb & "\Link_Tab.dsn") Then
            oIni2.SetKeyValue("ODBC", "DRIVER", "Driver da Microsoft para arquivos texto (*.txt; *.csv)")
            oIni2.SetKeyValue("ODBC", "UID", "admin")
            oIni2.SetKeyValue("ODBC", "UserCommitSync", "Yes")
            oIni2.SetKeyValue("ODBC", "Threads", "3")
            oIni2.SetKeyValue("ODBC", "SafeTransactions", "0")
            oIni2.SetKeyValue("ODBC", "PageTimeout", "5")
            oIni2.SetKeyValue("ODBC", "MaxScanRows", "50")
            oIni2.SetKeyValue("ODBC", "MaxBufferSize", "2048")
            oIni2.SetKeyValue("ODBC", "FIL", "Text")
            oIni2.SetKeyValue("ODBC", "Extensions", "txt,csv,tab,asc")
            oIni2.SetKeyValue("ODBC", "DriverId", "27")
        End If
        If igInd_Kar = 0 Then
            oIni2.SetKeyValue("ODBC", "DefaultDir", sgDir_Fls)
            oIni2.SetKeyValue("ODBC", "DBQ", sgDir_Fls)
        Else
            oIni2.SetKeyValue("ODBC", "DefaultDir", sgDir_Fls2)
            oIni2.SetKeyValue("ODBC", "DBQ", sgDir_Fls2)
        End If
        oIni2.Save(sgDir_odb & "\Link_Tab.dsn")

        '---------------------------------------Link_Dbf.dsn---------------------------------------
        oIni2.Load(sgDir_odb & "\Link_Dbf.dsn")
        If Not IO.File.Exists(sgDir_odb & "\Link_Dbf.dsn") Then
            oIni2.SetKeyValue("ODBC", "DRIVER", "Microsoft Visual FoxPro Driver")
            oIni2.SetKeyValue("ODBC", "UID", "")
            oIni2.SetKeyValue("ODBC", "Collate", "Machine")
            oIni2.SetKeyValue("ODBC", "BackgroundFetch", "Yes")
            oIni2.SetKeyValue("ODBC", "Exclusive", "No")
            oIni2.SetKeyValue("ODBC", "SourceType", "DBF")
        End If
        oIni2.SetKeyValue("ODBC", "SourceDB", sgDir_Tmp)
        oIni2.Save(sgDir_odb & "\Link_Dbf.dsn")

        '------------------------------------------------schema.ini---------------------------------------
        oIni2.Load(sgDir_odb & "\schema.ini")
        If Not IO.File.Exists(sgDir_Fls & "\schema.ini") Then
            '---------------------------------------[file05.tab]--------------------------------------
            oIni2.SetKeyValue("file05.tab", "ColNameHeader", "False")
            oIni2.SetKeyValue("file05.tab", "Format", "CSVDelimited")
            oIni2.SetKeyValue("file05.tab", "MaxScanRows", "50")
            oIni2.SetKeyValue("file05.tab", "CharacterSet", "OEM")
            oIni2.SetKeyValue("file05.tab", "Col1", "ID_CAN Integer")
            oIni2.SetKeyValue("file05.tab", "Col2", "ID_COD Char Width 65")
            oIni2.SetKeyValue("file05.tab", "Col3", "DE_CAN Char Width 50")
            oIni2.SetKeyValue("file05.tab", "Col4", "FL_MP3 Char Width 80")
            oIni2.SetKeyValue("file05.tab", "Col5", "FL_MP3 Char Width 80")

            If Not IO.File.Exists(sgDir_Fls & "\file05.tab") Then
                Dim ofile As System.IO.StreamWriter
                ofile = My.Computer.FileSystem.OpenTextFileWriter(sgDir_Fls & "\file05.tab", True)
                ofile.WriteLine("")
                ofile.Close()
            End If

            oIni2.Save(sgDir_odb & "\schema.ini")
        End If
        oIni2 = Nothing
    End Sub

    Private Sub Colocar_Frames()
        With Me.oGrpBox_Dis
            .Height = oGrpDisc_Height
            .Width = oGrpDisc_With
            .Left = oGrpDisc_x
            .Top = oGrpDisc_y
        End With

        With Me.oGrpBox_Can
            .Height = oGrpCanc_Height
            .Width = oGrpCanc_With
            .Left = oGrpCanc_x
            .Top = oGrpCanc_y
        End With
        With Me.oGrpBox_Gen
            .Height = oGrpGene_Height
            .Width = oGrpGene_With
            .Left = oGrpGene_x
            .Top = oGrpGene_y
        End With
    End Sub

    Private Sub Img_PagPrev_Click(sender As Object, e As EventArgs) Handles Img_PagPrev.Click
        Me.Change_Page_Up()
    End Sub

    Private Sub Img_PagNext_Click(sender As Object, e As EventArgs) Handles Img_PagNext.Click
        Me.Change_Page_Dn()
    End Sub

    Private Sub oTime_Mensajes_Tick(sender As Object, e As EventArgs) Handles oTime_Mensajes.Tick
        If Me.oTimer_Reset.Enabled = False Then
            Me.oTimer_Reset.Enabled = True
        End If
        If (Me.olMessage.BackColor) = Color.White Then
            Me.olMessage.BackColor = Color.Transparent

        Else
            Me.olMessage.BackColor = Color.White
        End If
        Application.DoEvents()
    End Sub

    Private Sub oTimer_Reset_Tick(sender As Object, e As EventArgs) Handles oTimer_Reset.Tick
        Me.Muestra_Tema_Det()
        Me.oTime_Mensajes.Enabled = False
        Me.olMessage.Visible = False
        Me.olMensaje_Video.Visible = False
        Application.DoEvents()
    End Sub

    Private Sub oGeneral_Timer_Tick(sender As Object, e As EventArgs) Handles oGeneral_Timer.Tick
        '------------------------------------------------------
        If bgBlinkPag = True Then
            If Me.olPaginas.Visible = True Then
                If (Me.olPaginas.ForeColor) <> ConvertToRbg(&HFFFF&) Then
                    Me.olPaginas.ForeColor = ConvertToRbg(&HFFFF&)
                Else
                    Me.olPaginas.ForeColor = ConvertToRbg(&H0&)
                End If
            End If
            If Me.Img_PagNext.Visible = True Then
                If Me.Img_PagNext.BorderStyle = 0 Then
                    Me.Img_PagNext.BorderStyle = 1
                Else
                    Me.Img_PagNext.BorderStyle = 0
                End If
            End If
            If Me.Img_PagPrev.Visible = True Then
                If Me.Img_PagPrev.BorderStyle = 0 Then
                    Me.Img_PagPrev.BorderStyle = 1
                Else
                    Me.Img_PagPrev.BorderStyle = 0
                End If
            End If
            If olPaginas2.Visible = True Then
                If (Me.olPaginas2.ForeColor) <> ConvertToRbg(&HFFFF&) Then
                    Me.olPaginas2.ForeColor = ConvertToRbg(&HFFFF&)
                Else
                    Me.olPaginas2.ForeColor = ConvertToRbg(&H0&)
                End If
            End If
        Else
            Me.olPaginas.ForeColor = ConvertToRbg(&HFFFF&)
            Me.olPaginas2.ForeColor = ConvertToRbg(&HFFFF&)
            Me.Img_PagPrev.BorderStyle = 0
            Me.Img_PagNext.BorderStyle = 0
        End If
        '------------------------------------------------------
        If ((Strings.Trim(Me.otCodigo.Text) = "") Or (igNext_Return_Gen = 0)) Then
            Exit Sub
        End If
        If igDelay_Return_Gen > 0 Then
            'If igCan_Sel = "" Then
            'Sólo regresa a género si esta en la pantalla de discos
            If igNext_Return_Gen <= (Hour(Now()) * 60) + Minute(Now()) Then
                Me.Retrocede()
                Me.Retrocede()
                igNext_Return_Gen = (Hour(Now()) * 60) + (Minute(Now()) + igDelay_Return_Gen)
            End If
            'End If
        End If

    End Sub

    Private Sub Retrocede()
        If Me.otCodigo.Focused() = False Then
            Me.otCodigo.Focus()
        End If
        Dim sValue As String = Trim(otCodigo.Text)
        igLen = Len(sValue)

        Select Case igLen
            Case Is = 1
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
            Case Is = 2
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
            Case Is = 3
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
            Case Is = 4
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
            Case Is = 5
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
            Case Is = 6
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
            Case Else
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
                My.Computer.Keyboard.SendKeys("{BACKSPACE}")
        End Select
    End Sub

    Private Sub oTM_codigo2_Tick(sender As Object, e As EventArgs) Handles oTM_codigo2.Tick
        Call Retrocede()

        If Mid(Trim(Me.otCodigo.Text), 1, 2) = "99" Then
            Call Retrocede()
            Call Retrocede()
        End If
        Me.oTM_codigo2.Enabled = False
        igNext_Return_Gen = (Hour(Now()) * 60) + (Minute(Now()) + igDelay_Return_Gen)
        Application.DoEvents()
    End Sub

    Private Sub oTimer_Moneda_Tick(sender As Object, e As EventArgs) Handles oTimer_Blink.Tick
        If igKeep_Cred = 1 Then
            Exit Sub
        End If
        If igCnt_CR > 0 Then
            Exit Sub
        End If
        If (Me.olCred_Msg.ForeColor) = Color.Yellow Then
            Me.olCred_Msg.ForeColor = Color.Black
        Else
            Me.olCred_Msg.ForeColor = Color.Yellow
        End If

        If Me.olMessage.ForeColor = Color.Yellow Then
            Me.olMessage.ForeColor = Color.Black
        Else
            Me.olMessage.ForeColor = Color.Yellow
        End If

        If Me.oLTitulo.ForeColor = Color.Black Then
            Me.oLTitulo.ForeColor = Color.Yellow
        Else
            Me.oLTitulo.ForeColor = Color.Black
        End If
        If Me.otCodigo.Focused = False Then
            Me.otCodigo.Focus()
        End If
        Application.DoEvents()
    End Sub

    Private Sub otCodigo_TextChanged(sender As Object, e As EventArgs) Handles otCodigo.TextChanged
        REM On Error GoTo Solve_error
        Dim sValue As String = Me.otCodigo.Text.Trim().Replace("-", "")
        igLen = Trim(sValue).Length


        If gbServ_Mode = True Then
            If Me.otCodigo.Mask = "##-##" Then
                Me.otCodigo.Mask = "##-##-##"
                Me.otCodigo.Tag = "0"
            End If
        Else
            If igLen > 0 Then
                REM Call Show_Hide_Service(0)
                If igKeep_Cred = 0 Then
                    If igCnt_CR <= 0 Then
                        igCnt_CR = 0
                        If Me.otCodigo.Tag = "" Then
                            Me.otCodigo.Mask = "##-##"
                            Me.otCodigo.Tag = "1"
                            Me.oSetFocus_Codigo.Focus()
                            Exit Sub
                        End If
                    Else
                        If Me.otCodigo.Mask = "##-##" Then
                            Me.otCodigo.Mask = "##-##-##"
                            Me.otCodigo.Tag = "0"
                        End If
                    End If
                Else
                    If Me.otCodigo.Mask = "##-##" Then
                        Me.otCodigo.Mask = "##-##-##"
                        Me.otCodigo.Tag = "0"
                    End If
                End If
            Else
                igGen_Sel = "" : igDis_Sel = "" : igCan_Sel = ""
            End If
        End If
        '   *******************************************************************************

        If Trim(Me.otCodigo.Text) = "99" Then
            If igShowPass = 1 Then
                If gbServ_Mode = False Then
                    Me.Tag = ""
                    Me.otCodigo.Text = ""
                    REM Call AlwaysOnTop(Main_Form, False)
                    REM Pass_Scr.Show vbModal
                    REM Call AlwaysOnTop(Main_Form, bgKeep_On_Top)
                    If Trim(Me.Tag) <> sgCr_AKey Then
                        Me.olMessage.Text = "ACCESO DENEGADO!!!"
                        Me.olMessage.Visible = True
                        Me.oTime_Mensajes.Enabled = True
                        Me.olMessage.Tag = ""
                        Me.oSetFocus_Codigo.Focus()
                        gbServ_Mode = False
                        Me.oService_Info.Clear()
                        Me.oService_Info.Visible = False
                        Exit Sub
                    End If
                    Me.olMessage.Text = "ACCESO CONSEDIDO!"
                    Me.olMessage.Tag = ""
                    Me.olMessage.Visible = True
                    pTimeResto = (2 * 30)
                    gbServ_Mode = True
                    Me.Timer3.Enabled = True
                    Me.oService_Info.LoadFile(Application.StartupPath & "\Service_Info.RTF")
                    Me.olT_Mant.Visible = True
                    Me.oService_Info.Top = 2160
                    Me.oService_Info.Width = 6735
                    Me.oService_Info.Left = 5760
                    Me.oService_Info.Height = 4455
                    Me.oService_Info.Visible = True
                End If
            End If
        End If
        '   *******************************************************************************
        Select Case igLen
            Case Is = 0
                Me.olTimer1.Text = 0
                Me.olTimer2.Text = 0

                Me.olTimer1.Visible = False
                Me.olTimer2.Visible = False

                igGen_Sel = ""
                igDis_Sel = ""
                igCan_Sel = ""

                igAct_PgG = 1
                igAct_PgD = 1
                igAct_PgC = 1

                igNext_Return_Gen = 0
                Me.Image2.Visible = False
                If Me.oGrpBox_Gen.Visible = False Then
                    Me.Cargar_Pag_Gen(1)
                End If
            Case 1 To 2
                If gbServ_Mode = True Then
                    Me.Desactiva_Cancion(True)
                    Me.Desactiva_Disco(True)
                    Me.Desactiva_Genero(True)
                    Me.oService_Info.LoadFile(Application.StartupPath & "\Service_Info.RTF")
                    Me.olT_Mant.Visible = True
                    Me.oService_Info.Visible = True
                    Exit Sub
                End If
                Dim sGen_Ret As String
                igGen_Sel = ""
                igDis_Sel = ""
                igCan_Sel = ""

                Me.Image2.ImageLocation = ""
                If (igLen < 2) Then
                    Exit Sub
                End If
                If (Me.Busca_Gen_Sel(sValue) = False) Then
                    If igLen = 2 Then
                        Me.oLTitulo.Text = "[GENERO NO EXISTE!!!]"
                    Else
                        Me.oLTitulo.Text = ""
                    End If
                    Me.Retrocede()
                    igAct_PgG = 1
                    Exit Sub
                Else
                    Me.oLTitulo.Text = "[Seleccione le Disco]..."
                End If
                igAct_PgD = 1
                igAct_PgC = 1
                '***********************DISCOS************************
                If Me.oGrpBox_Dis.Visible = False Then
                    If igLen = 2 Then
                        Me.Cargar_Pag_Dis(sgGen_Sel_Id, igAct_PgD)
                    End If
                End If
                    bgVIP = False
                igNext_Return_Gen = (Hour(Now()) * 60) + (Minute(Now()) + igDelay_Return_Gen)
                If (Conversion.Val(Me.olTimer1.Text) = 0) Then
                    Me.olTimer1.Text = Format(igDelay_Return_Gen * 60, "#0")
                    Me.olTimer1.Visible = True
                End If
            Case 3 To 4
                If (gbServ_Mode = True) Then
                    Exit Sub
                End If

                igDis_Sel = ""
                igCan_Sel = ""
                Me.Image2.ImageLocation = Nothing
                If (igLen < 4) Then
                    REM Me.Desactiva_Disco(False)
                    Exit Sub
                End If
                Dim sValSel As String = Strings.Mid(sValue, 3, 4)
                If Me.Busca_Sel_Dis(sValSel) = False Then
                    If igLen >= 4 Then
                        Me.oLTitulo.Text = "Disco no Existe..."
                    Else
                        Me.oLTitulo.Text = "[Seleccione le Disco]..."
                    End If
                    Me.Retrocede()
                    Exit Sub
                Else
                    Me.oLTitulo.Text = "[Seleccione le Tema]..."
                End If
                igAct_PgC = 1

                Me.Image2.Visible = True
                '***********************CANCION************************
                If Me.oGrpBox_Can.Visible = False Then
                    If igLen = 4 Then
                        Me.Cargar_Pag_Canc(sgDis_Sel_Id, igAct_PgC)
                    End If
                End If

                    bgVIP = False
            Case 5 To 6
                If gbServ_Mode = True Then
                    Select Case Strings.Trim(Me.otCodigo.Text)
                        Case Is = "990001"
                            My.Computer.Keyboard.SendKeys(sgKb_SwP)
                            My.Computer.Keyboard.SendKeys(".")
                            My.Computer.Keyboard.SendKeys(".")
                            My.Computer.Keyboard.SendKeys(".")
                        Case Is = "990002"
                            My.Computer.Keyboard.SendKeys(sgKb_Crd2)
                        Case Is = "990003"
                            My.Computer.Keyboard.SendKeys(sgKb_Del)
                            My.Computer.Keyboard.SendKeys(sgKb_Del)
                            My.Computer.Keyboard.SendKeys(sgKb_Del)
                        Case Is = "990004"
                            My.Computer.Keyboard.SendKeys(sgKb_ResM)
                        Case Is = "990005"
                            My.Computer.Keyboard.SendKeys(sgKb_ResA)
                        Case Is = "990006"
                            pTimeResto = 1
                        Case Is = "990007"
                            If (IO.File.Exists(Application.StartupPath & "\osk.exe") = True) Then
                                Shell(Application.StartupPath & "\osk.exe")
                            End If
                        Case Is = "990008"
                            REM Call Show_Hide_Service
                        Case Is = "990009"
                            My.Computer.Keyboard.SendKeys(sgKb_SwK)
                        Case Is = "990010"
                            REM Call Go_Service
                        Case Is = "990011"
                            If (IO.File.Exists(Application.StartupPath & "\TSR.exe") = True) Then
                                Shell(Application.StartupPath & "\TSR.exe")
                            End If
                        Case Is = "990012"
                            Shell("EXPLORER.EXE")
                        Case Is = "990013"
                            REM Unload Video_Form
                            End
                        Case Is = "990014"
                            If (igFlg_TCR = 1) Then
                                igFlg_TCR = 0
                            Else
                                igFlg_TCR = 1
                            End If
                            oIni.SetKeyValue("CREDITS", "FLG_TESTCR", Strings.Format(igFlg_TCR, "#####0"))
                            oIni.Save(cini_path)
                        Case Else
                            If (igLen = 6) Then
                                Me.otCodigo.Text = ""
                            End If
                    End Select
                    Me.olTimer2.Text = Format(Me.oTM_codigo2.Interval / 1000, "###0")
                    Me.oTM_codigo2.Enabled = True

                    Me.olMessage.Text = ""
                    Me.olMessage.Tag = "1"
                    Me.olMessage.Visible = True
                    Me.oTimer_Srv.Enabled = True
                    Exit Sub
                End If

                Dim sCad_Ato As String
                Dim sFle_Mp3 As String
                igCan_Sel = ""
                If (igLen < 6) Then
                    Exit Sub
                End If
                Dim sValSel As String = Strings.Mid(sValue, 5, 6)
                If Me.Busca_Sel_Canc(sValSel, sCad_Ato, sFle_Mp3) = False Then
                    If (igLen >= 6) Then
                        Me.oLTitulo.Text = "Tema no existe..."
                    Else
                        Me.oLTitulo.Text = ""
                    End If
                    igAct_PgC = 1
                    Me.Retrocede()
                    Exit Sub
                Else
                    Me.oLTitulo.Text = "[Seleccione le Tema]..."
                End If
                If Me.oGrpBox_Can.Visible = False Then
                    Me.Cargar_Pag_Canc(igDis_Sel, igAct_PgC)
                End If

                If igLen = 6 Then
                    If IO.File.Exists(sFle_Mp3) Then
                        If bgVIP = False Then
                            If Not Strings.UCase(Strings.Right(Strings.Trim(sFle_Mp3), 3)) <> "MP3" Then
                                If igNoDuplicT = 1 Then
                                    If Me.Busca_ATocar(sCad_Ato) = False Then
                                        Me.oLst_A_Tocar.Items.Add(sCad_Ato)
                                    End If
                                Else
                                    Me.oLst_A_Tocar.Items.Add(sCad_Ato)
                                End If
                                If igKeep_Cred = 0 Then
                                    If igCnt_CR > 0 Then
                                        igCnt_CR = igCnt_CR - 1
                                    End If
                                End If
                            Else
                                If sgKb_VID = 0 Then
                                    If igNoDuplicT = 1 Then
                                        If Me.Busca_ATocar(sCad_Ato) = False Then
                                            Me.oLst_A_Tocar.Items.Add(sCad_Ato)
                                        End If
                                    Else
                                        Me.oLst_A_Tocar.Items.Add(sCad_Ato)
                                    End If
                                    If igKeep_Cred = 0 Then
                                        If igCnt_CR > 0 Then
                                            igCnt_CR = igCnt_CR - 1
                                        End If
                                    End If
                                Else
                                    If sgKb_VID > igCnt_CR Then
                                        Me.olMessage.Visible = True
                                        Me.olMessage.Text = "CREDITOS INSUFICIENTES!"
                                        Me.oTime_Mensajes.Enabled = True
                                        Call Retrocede()
                                        Exit Sub
                                    Else
                                        If igKeep_Cred = 0 Then
                                            igCnt_CR = igCnt_CR - sgKb_VID
                                        End If
                                        If igNoDuplicT = 1 Then
                                            If Me.Busca_ATocar(sCad_Ato) = False Then
                                                Me.oLst_A_Tocar.Items.Add(sCad_Ato)
                                            End If
                                        Else
                                            Me.oLst_A_Tocar.Items.Add(sCad_Ato)
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If Me.oLst_A_Tocar.Items.Count Then
                                Me.oLst_A_Tocar.Items.Add(sCad_Ato)
                            Else
                                Me.oLst_A_Tocar.Items.Insert(1, sCad_Ato)
                            End If
                            If igKeep_Cred = 0 Then
                                igCnt_CR = igCnt_CR - 2
                            End If
                            If igCnt_CR < 0 Then
                                igCnt_CR = 0
                            End If
                            bgVIP = False
                            Me.olMessageVIP.Visible = False
                        End If
                        '---Segmento que añade los bonus de promociones--
                        If sgIdx_Prm > 0 Then
                            piCnt_Canc = piCnt_Canc + 1
                            If piCnt_Canc >= sgIdx_Prm Then
                                piCnt_Canc = 0
                                Dim sCad_Ato2 As String = Me.Add_Promos()

                                If Not (Strings.Trim(sCad_Ato2) = "") Then
                                    Me.oLst_A_Tocar.Items.Add(sCad_Ato2)
                                    Me.olMessage.Visible = True
                                    Me.olMessage.Text = "DISCO PROMO EN COLA!"
                                    Me.oTime_Mensajes.Enabled = True

                                End If
                            End If
                        End If
                        '------------------------------------------------
                        Me.oLst_A_Tocar.Refresh()
                        Me.Refresh_Creditos()
                    Else
                        Me.olMessage.Visible = True
                        Me.olMessage.Text = "NO ENCONTRADO!"
                        Me.oTime_Mensajes.Enabled = True
                        Me.Retrocede()
                        Exit Sub
                    End If
                    If igKeep_Cred = 0 Then
                        If igCnt_CR > 0 Then
                            Me.olMessage.Visible = True
                            Me.olMessage.Text = "TEMA FUE ANEXADO!"
                            Me.oTime_Mensajes.Enabled = True
                            'Call Retrocede
                            Me.olTimer2.Text = Format(Me.oTM_codigo2.Interval / 1000, "###0")
                            Me.oTM_codigo2.Enabled = True
                        Else
                            '----------------------GENEROS------------------------------
                            Me.Cargar_Pag_Gen(1)
                            Me.oSetFocus_Codigo.Focus()
                        End If
                    Else
                        Me.olMessage.Visible = True
                        Me.olMessage.Text = "TEMA FUE ANEXADO!"
                        Me.oTime_Mensajes.Enabled = True
                        Me.olTimer2.Text = Format(Me.oTM_codigo2.Interval / 1000, "###0")
                        Me.oTM_codigo2.Enabled = True
                    End If
                    Me.Salvar_Temas()
                End If
                Me.olTimer2.Visible = True

        End Select
    End Sub

    Private Function Busca_ATocar(pCad_Ato As String) As Boolean
        Dim i As Integer
        Dim iCont As Integer
        iCont = 0
        Busca_ATocar = False
        For i = 1 To Me.oLst_A_Tocar.Items.Count - 1
            If Me.oLst_A_Tocar.Items(i) = pCad_Ato Then
                Busca_ATocar = True
                Exit Function
            End If
            'VBA.DoEvents
        Next i
    End Function

    Private Function Busca_Gen_Sel(ByVal pValor_Bus As String) As Boolean
        Dim dar As SQLiteDataReader
        Dim sql As String
        '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------     

        If Not (oConnSqLite.State = ConnectionState.Open) Then
            oConnSqLite.Open()
        End If

        sql = "SELECT * FROM file01 WHERE id_ord='" & Strings.Trim(pValor_Bus) & "' LIMIT 1"
        oCmmSqLite = New SQLiteCommand(sql, oConnSqLite)
        dar = oCmmSqLite.ExecuteReader()

        '----------------------------------------------------------------------------------
        Try
            If dar.HasRows = True Then
                dar.Read()
                sgGen_Sel_Id = dar.Item("id_gen").ToString()
                igGen_Sel = dar.Item("id_ord").ToString()
                igAct_PgG = Conversion.Val(dar.Item("page").ToString())
                Return True
            Else
                sgGen_Sel_Id = ""
                igGen_Sel = ""
                igAct_PgG = 0
                Return False
            End If
            '----------------------------------------------------------------------------------

        Catch ex As SQLiteException
            Console.WriteLine("Error: " & ex.ToString())
            sgGen_Sel_Id = ""
            igGen_Sel = ""
            igAct_PgG = 0
            Return False
        Finally
            dar.Close()
            dar = Nothing
        End Try
    End Function

    Private Function Busca_Sel_Dis(ByVal pValor_Bus As String) As Boolean
        Dim dar As SQLiteDataReader
        Dim sql As String
        '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------

        If Not (oConnSqLite.State = ConnectionState.Open) Then
            oConnSqLite.Open()
        End If

        sql = "SELECT * FROM file02 WHERE id_gen='" & Strings.Trim(sgGen_Sel_Id) & "' AND id_ord='" & Strings.Trim(pValor_Bus) & "' LIMIT 1"
        oCmmSqLite = New SQLiteCommand(sql, oConnSqLite)
        dar = oCmmSqLite.ExecuteReader()

        '----------------------------------------------------------------------------------
        Try
            If dar.HasRows = True Then
                dar.Read()
                sgDis_Sel_Id = dar.Item("id_dis").ToString()
                igDis_Sel = dar.Item("id_ord").ToString()
                igAct_PgD = Conversion.Val(dar.Item("page").ToString())
                Dim sNamFl As String = System.IO.Path.GetFileName(Strings.Trim(dar.Item("fl_img").ToString))
                Dim sFile As String = sgDir_Img & sNamFl
                Me.Image2.ImageLocation = sFile
                Return True
            Else
                Me.Image2.ImageLocation = Nothing
                sgDis_Sel_Id = ""
                igDis_Sel = ""
                igAct_PgD = 0
                Return False
            End If
            '----------------------------------------------------------------------------------

        Catch ex As SQLiteException
            Console.WriteLine("Error: " & ex.ToString())
            Me.Image2.ImageLocation = Nothing
            sgDis_Sel_Id = ""
            igDis_Sel = ""
            igAct_PgD = 0
            Return False
        Finally
            dar.Close()
            dar = Nothing
        End Try
    End Function

    Private Function Busca_Sel_Canc(ByVal pValor_Bus As String, ByRef cLineAdd As String, ByRef cFilemp3 As String) As Boolean

        Dim dar As SQLiteDataReader
        Dim sql As String
        '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------

        If Not (oConnSqLite.State = ConnectionState.Open) Then
            oConnSqLite.Open()
        End If

        sql = "SELECT " &
            "file03.id_ord," &
            "file03.ID_CAN," &
            "file03.DE_CAN," &
            "file03.FL_MP3," &
            "file03.page," &
             "file02.fl_img " &
            "FROM file03 " &
            "LEFT JOIN file02 ON file03.id_dis = file02.id_dis " &
            "WHERE file03.id_dis = '" & Strings.Trim(sgDis_Sel_Id) & "' And file03.id_ord = '" & Strings.Trim(pValor_Bus) & "' LIMIT 1"

        oCmmSqLite = New SQLiteCommand(sql, oConnSqLite)
        dar = oCmmSqLite.ExecuteReader()

        '----------------------------------------------------------------------------------
        Try
            If dar.HasRows = True Then
                dar.Read()
                sgCan_Sel_Id = dar.Item("id_can").ToString()
                igCan_Sel = dar.Item("id_ord").ToString()
                igAct_PgC = Conversion.Val(dar.Item("page").ToString())

                Dim sFld1 As String = dar.Item("id_ord").ToString()
                Dim sFld2 As String = dar.Item("ID_CAN").ToString()
                Dim sFld3 As String = dar.Item("DE_CAN").ToString()
                Dim sNamFl As String = System.IO.Path.GetFileName(Strings.Trim(dar.Item("FL_MP3").ToString()))
                Dim sFld4 As String = sgDir_Mp3 & sNamFl
                Dim sFld5 As String = ""
                Dim sFld6 As String = dar.Item("fl_img").ToString()
                cLineAdd = sFld1 & "," & sFld2 & "," & sFld3 & "," & sFld4 & "," & sFld5 & "," & sFld6
                cFilemp3 = sFld4
                Return True
            Else
                Me.Image2.ImageLocation = Nothing
                sgCan_Sel_Id = ""
                igCan_Sel = ""
                igAct_PgC = 0
                cLineAdd = ""
                Return False
            End If
            '----------------------------------------------------------------------------------

        Catch ex As SQLiteException
            Console.WriteLine("Error: " & ex.ToString())
            sgCan_Sel_Id = ""
            igCan_Sel = ""
            igAct_PgC = 0
            cLineAdd = ""
            Return False
        Finally
            dar.Close()
            dar = Nothing
        End Try
    End Function

    Private Sub oSetFocus_Codigo_Enter(sender As Object, e As EventArgs) Handles oSetFocus_Codigo.Enter
        Me.otCodigo.Text = ""
        Me.otCodigo.Focus()
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Dim Horas As Integer, Minutos As Integer, Segundos As Integer, Cadena As String

        Segundos = pTimeResto
        Horas = Int(Segundos / 3600)
        Segundos = Segundos Mod 3600
        Minutos = Int(Segundos / 60)
        Segundos = Segundos Mod 60
        Cadena = Format$(Horas, "00") & ":" & Format$(Minutos, "00") & ":" & Format$(Segundos, "00") & " Restante"
        Me.olT_Mant.Text = "Opción de servicio activada" & Chr(13) & "[" & Cadena & "] Restante " & Chr(13) & "99-00-06 DESACTIVAR"
        pTimeResto = pTimeResto - 1
        If pTimeResto = -1 Then
            gbServ_Mode = False
            Me.Timer3.Enabled = False
            Me.oService_Info.Clear()
            Me.olT_Mant.Visible = False
            Me.oService_Info.Visible = False

            Me.Retrocede()
            Me.Retrocede()
        End If
    End Sub

    Private Sub otCodigo_KeyDown(sender As Object, e As KeyEventArgs) Handles otCodigo.KeyDown
        If (e.KeyCode = 888888) Then
            Dim sValue As String = Me.otCodigo.Text.Trim().Replace("-", "")
            igLen = Trim(sValue).Length
            Select Case igLen
                Case <= 2
                    If Me.oGrpBox_Gen.Visible = False Then
                        If Me.oGrpBox_Can.Visible = True Then
                            Me.Desactiva_Cancion(True)
                        End If
                        If Me.oGrpBox_Dis.Visible = True Then
                            Me.Desactiva_Disco(True)
                        End If
                        Me.Desactiva_Genero(False)
                    End If
                Case <= 4
                    If Me.oGrpBox_Dis.Visible = False Then
                        If Me.oGrpBox_Can.Visible = True Then
                            Me.Desactiva_Cancion(True)
                        End If
                        If Me.oGrpBox_Gen.Visible = True Then
                            Me.Desactiva_Genero(True)
                        End If
                        Me.Desactiva_Genero(False)
                    End If
            End Select

        End If

        If (e.KeyCode = 123) Then 'F12 PARA SALIR DEL SISTEMA
            REM Call Form_Unload(1)
            REM Unload Video_Form
            Console.WriteLine("F12 PARA SALIR DEL SISTEMA")
            End
        End If
        If (e.KeyCode = 122) Then 'F11 PARA MENÚ DE SERVICIO
            REM Call Show_Hide_Service
            Console.WriteLine("F11 PARA MENÚ DE SERVICIO")
        End If
    End Sub

    Private Sub otCodigo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles otCodigo.KeyPress
        Dim iLimCnt As Integer
        igKeyAscii = Asc(e.KeyChar)
        REM If IsNumeric(Chr(KeyAscii)) Then
        REM Exit Sub
        REM End If

        If Asc(e.KeyChar) = 8 Then
            Me.Focus()
            Exit Sub
        End If
        If (UCase(e.KeyChar) = sgKb_SwK) Then
            igInd_Kar = IIf(igInd_Kar = 0, 1, 0)
            oIni.SetKeyValue("GENERAL", "SWITCH_KAR", Format(igInd_Kar, "#####0"))
            oIni.SetKeyValue("GENERAL", "RELOAD_APP", "1")
            Shell(Application.StartupPath & "\R_monitor.exe", AppWinStyle.NormalFocus)
            End
            Exit Sub
        End If
        If (UCase(e.KeyChar) = sgKb_Pause) Then
            If Me.oLst_A_Tocar.Items.Count = 0 Then
                Exit Sub
            End If
            'If igScr_Alone = 1 Then
            If bgIs_Video = True Then
                If Me.MediaPlayer2.playState = WMPLib.WMPPlayState.wmppsPaused Then
                    Me.MediaPlayer2.Ctlcontrols.play()
                    Video_Form.MediaPlayer3.Ctlcontrols.play()
                    Me.otCargador_Music.Enabled = True
                Else
                    Me.MediaPlayer2.Ctlcontrols.pause()
                    Video_Form.MediaPlayer3.Ctlcontrols.pause()
                    Me.otCargador_Music.Enabled = False
                End If
            End If
            'Else
            'End If
            Exit Sub
        End If
        '133
        If (UCase(e.KeyChar) = sgKb_Vef) Then
            bgExit = False
            REM bFlagChek = False
            If TBack4.Visible = False Then
                TBack4.Visible = True
            Else
                TBack4.Visible = False
            End If
            e.KeyChar = ChrW(Keys.None)
        End If

        If (UCase(e.KeyChar) = sgKb_SwP) Then
            Me.olMessage.Visible = True
            Me.olMessage.Text = "SWITCHING VIDEO..."
            Me.oTime_Mensajes.Enabled = True
            igInd_Pub = 0
            If bgSw_Pub = False Then
                Me.oInd_VideoSW.Image = My.Resources.grnled
                Me.oInd_VideoSW.Visible = True
            Else
                'Me.oInd_VideoSW.Picture = LoadPicture(App.Path & "\" & "grnled.GIF")
                Me.oInd_VideoSW.Visible = False
            End If
            bgSw_Pub = IIf(bgSw_Pub = True, False, True)
            REM Me.Conectar_DBPub()
            e.KeyChar = ChrW(Keys.None)
        End If

        If (UCase(e.KeyChar) = sgKb_Ret) Then
            Call Retrocede()
            e.KeyChar = ChrW(Keys.None)
        End If

        If (UCase(e.KeyChar) = sgKb_ResM) Then
            'Sección que se ejecuta si se preciona [S/s] (Reset single active music)
            Video_Form.MediaPlayer3.close()
            Me.MediaPlayer2.close()
            REM Main_Form.MediaPlayer1.Close

            Me.oLDuracion.Text = "00:00"
            Me.olAct_Pos.Text = "00:00"
            bfTm = False

            If igInd_Pub < igTot_Pub Then
                igInd_Pub = igInd_Pub + 1
            Else
                igInd_Pub = 0
            End If
            Me.olMessage.Visible = True
            Me.olMessage.Text = "TEMA IGNORADO"
            Me.oTime_Mensajes.Enabled = True
            If Me.oLst_A_Tocar.Items.Count > 0 Then
                Me.oLst_A_Tocar.Items.RemoveAt(0)
            Else
                Me.olMessage.Visible = True
                Me.olMessage.Text = "TEMAS AGOTADOS"
                Me.oTime_Mensajes.Enabled = True
            End If
            Me.Cargar_Gen()
            Me.Muestra_Tema_Det()
            Me.Refresh_Creditos()
            bgWMP_Busy = False
            igCont_Sin = 0
            Me.Salvar_Temas()
            Me.oSetFocus_Codigo.Focus()
            e.KeyChar = ChrW(Keys.None)
        End If

        If (UCase(e.KeyChar) = sgKb_ResA) Then
            'Sección que se ejecuta si se preciona [R/r] (Resert all)

            Me.oLDuracion.Text = "00:00"
            Me.olAct_Pos.Text = "00:00"
            bfTm = False
            Me.olMessage.Visible = True
            Me.olMessage.Text = "CREDITOS ANULADOS"
            Me.oTime_Mensajes.Enabled = True

            Me.oLst_A_Tocar.Items.Clear()
            Video_Form.MediaPlayer3.close()
            Me.MediaPlayer2.close()
            REM Main_Form.MediaPlayer1.Close
            Me.oSetFocus_Codigo.Focus()

            If igInd_Pub < igTot_Pub Then
                igInd_Pub = igInd_Pub + 1
            Else
                igInd_Pub = 0
            End If
            Me.Cargar_Pag_Gen(1)

            Me.Muestra_Tema_Det()
            Me.Refresh_Creditos()
            bgWMP_Busy = False
            igCont_Sin = 0
            If (igKeep_Cred = 0) Then
                igCnt_CR = 0
                If igFlg_SavedCR = 1 Then
                    oIni.SetKeyValue("CREDITS", "ACU_SAVECR", Format(igCnt_CR, "#####0"))
                End If
            Else
                igCnt_CR = 6
            End If
            Me.Salvar_Temas()
            Me.oSetFocus_Codigo.Focus()
            e.KeyChar = ChrW(Keys.None)
        End If

        If (UCase(e.KeyChar) = sgKb_Crd1) Then
            Me.Add_CR1()
            Me.Add_CR2()
            Me.Salvar_Temas()
            e.KeyChar = ChrW(Keys.None)
        End If

        If (UCase(e.KeyChar) = sgKb_Del) Then
            'Sección que se ejecuta si se preciona [-] (Crédito)
            If igKeep_Cred = 0 Then
                If igCnt_CR > 0 Then
                    igCnt_CR = igCnt_CR - 1
                    Me.Refresh_Creditos()
                End If
            End If
            e.KeyChar = ChrW(Keys.None)
        End If

        If (UCase(e.KeyChar) = sgKb_Crd2) Then
            If Strings.Trim(sgCr_AKey) = "" Then
                Me.olMessage.Text = "NO SE HA CONDIGURADO ACCESO"
                Me.olMessage.Visible = True
                Me.oTime_Mensajes.Enabled = True
                Me.oSetFocus_Codigo.Focus()
                e.KeyChar = ChrW(Keys.None)
            End If
            igCnt_CR = igCnt_CR + 6
            Me.Refresh_Creditos()
            Me.otCodigo.Mask = "##-##-##"
            Me.otCodigo.Tag = ""
            Me.oSetFocus_Codigo.Focus()
            e.KeyChar = ChrW(Keys.None)
        End If

        If (UCase(e.KeyChar) = sgKb_Pop) Then
            'Sección que se ejecuta si se preciona [P/p] (Popular)
            If igLen < 4 Then
                Exit Sub
            End If
            If (igLen > 3) And (igLen < 7) Then
                If igKeep_Cred = 0 Then
                    If (igCnt_CR < Me.oLst_Popular.Items.Count) Then
                        Me.olMessage.Text = "CRÉDITOS INSUFICIENTES"
                        Me.olMessage.Visible = True
                        Me.oTime_Mensajes.Enabled = True
                    Else
                        Me.olMessage.Text = "CARGANDO POPULAR"
                        Me.olMessage.Visible = True
                        Me.oTime_Mensajes.Enabled = True
                        For iLimCnt = 1 To (Me.oLst_Popular.Items.Count)
                            igCnt_CR = igCnt_CR - 1
                            Me.Refresh_Creditos()
                            Me.oLst_A_Tocar.Items.Add(Me.oLst_Popular.Items(0))
                            Me.oLst_Popular.Items.RemoveAt(0)
                        Next iLimCnt
                        If igMixe_Popu > 0 Then
                            Call Shuffle_ListBox(Me.oLst_A_Tocar)
                        End If
                    End If
                Else
                    Me.olMessage.Text = "CARGANDO POPULAR"
                    Me.olMessage.Visible = True
                    Me.oTime_Mensajes.Enabled = True
                    For iLimCnt = 1 To (Me.oLst_Popular.Items.Count)
                        If igKeep_Cred = 0 Then
                            igCnt_CR = igCnt_CR - 1
                        End If
                        Me.Refresh_Creditos()
                        Me.oLst_A_Tocar.Items.Add(Me.oLst_Popular.Items(0))
                        Me.oLst_Popular.Items.RemoveAt(0)
                    Next iLimCnt
                    If igMixe_Popu > 0 Then
                        Call Shuffle_ListBox(Me.oLst_A_Tocar)
                    End If
                End If
                Call Cargar_Gen()
                Me.oSetFocus_Codigo.Focus()
                e.KeyChar = ChrW(Keys.None)
            End If
        End If

        If (UCase(e.KeyChar) = sgKb_VIP) Then
            'Sección que se ejecuta si se preciona [V/v] (VIP)
            If igLen < 4 Then
                Exit Sub
            End If
            If (igLen > 3) And (igLen < 7) Then
                If igKeep_Cred = 0 Then
                    If (igCnt_CR < 2) Then
                        Me.olMessage.Text = "CRÉDITOS INSUFICIENTES"
                        Me.olMessage.Visible = True
                        Me.oTime_Mensajes.Enabled = True
                        bgVIP = False
                    Else
                        Me.olMessageVIP.Text = "VIP EN PROCESO"
                        Me.olMessageVIP.Visible = True
                        bgVIP = True
                    End If
                Else
                    Me.olMessageVIP.Text = "VIP EN PROCESO"
                    Me.olMessageVIP.Visible = True
                    bgVIP = True
                End If
                e.KeyChar = ChrW(Keys.None)
            End If
        End If

        Select Case igLen
            Case 0 To 1
                Me.oGrpBox_Dis.Visible = False
                Me.oGrpBox_Can.Visible = False
                Me.Desplazar_Pantalla_Gen(e.KeyChar)
            Case 2 To 3
                Me.oGrpBox_Gen.Visible = False
                Me.oGrpBox_Can.Visible = False
                Me.Desplazar_Pantalla_Dis(e.KeyChar)

                igNext_Return_Gen = (Hour(Now()) * 60) + (Minute(Now()) + igDelay_Return_Gen)
                Me.olTimer1.Text = Format(igDelay_Return_Gen * 60, "#0")
                Me.olTimer1.Visible = True
            Case 4 To 6
                Me.oGrpBox_Gen.Visible = False
                Me.oGrpBox_Dis.Visible = False
                Dim bStatus As Boolean = Me.oGrpBox_Can.Visible
                Me.oGrpBox_Can.Visible = False
                Me.Desplazar_Pantalla_Can(e.KeyChar)
                Me.oGrpBox_Can.Visible = bStatus

                igNext_Return_Gen = (Hour(Now()) * 60) + (Minute(Now()) + igDelay_Return_Gen)
                Me.olTimer1.Text = Format(igDelay_Return_Gen * 60, "#0")
                Me.olTimer1.Visible = True
        End Select

    End Sub

    Public Sub Refresh_Creditos()
        If igKeep_Cred = 1 Then
            Me.olCreditos.Text = "CRÉDITOS GRATIS"
            Me.oTimer_Blink.Enabled = False
            Me.olCred_Msg.Visible = False
            Exit Sub
        End If
        If igCnt_CR > 0 Then
            Me.oTimer_Blink.Enabled = False
            Me.olCred_Msg.Visible = False
        Else
            Me.oTimer_Blink.Enabled = True
            Me.olCred_Msg.Visible = True
        End If
        Me.olCreditos.Text = "CREDITOS (" + Strings.Trim(Str(igCnt_CR)) & ")"
        Me.olMetros.Text = PADL(igCnt_CRG, 6, "0")
    End Sub

    Public Sub Add_CR1()
        If igKeep_Cred = 1 Then
            Exit Sub
        End If
        If (igFlg_TCR = 1) Then
            igCnt_CRP = igCnt_CRP + 1
            Me.olMetros2.Visible = True
            Me.olMetros2.Text = PADL(igCnt_CRP, 6, "0")
            Me.olLabMetros2.Visible = True
            Exit Sub
        Else
            Me.olMetros2.Visible = False
            Me.olLabMetros2.Visible = False
            igCnt_CRP = 0
        End If
        If (igCnt_CR >= igLim_Cred) Then
            Me.olMessage.Visible = True
            Me.olMessage.Text = "REVISAR MONEDERO!"
            Me.oTime_Mensajes.Enabled = True
        Else
            igCnt_CR = igCnt_CR + 2
            igCnt_CRG = igCnt_CRG + 1
            If (igCnt_CRG > 999999) Then
                oIni.SetKeyValue("CREDITS", "CEDIT_ACAN", Scramble(Format(igCnt_CRG, "#####0")))
                igCnt_CRG = 0
                oIni.SetKeyValue("CREDITS", "CEDIT_ACAC", Scramble(Format(igCnt_CRG, "#####0")))
                igCnt_CRG = Val(UnScramble(oIni.GetKeyValue("CREDITS", "CEDIT_ACAC", "0")))
            Else
                oIni.SetKeyValue("CREDITS", "CEDIT_ACAC", Scramble(Format(igCnt_CRG, "#####0")))
            End If
            If (igFlg_SavedCR = 1) Then
                oIni.SetKeyValue("CREDITS", "ACU_SAVECR", Format(igCnt_CR, "#####0"))
            End If
            oIni.Save(cini_path)

        End If
    End Sub

    Public Sub Add_CR2()
        If (sgKb_BonC > 0) Then
            If igCnt_CR = 8 Then
                Me.olMessage.Visible = True
                Me.olMessage.Text = "CRED. PROMOSIÓN [" & Strings.Trim(Str(sgKb_BonC)) & "]"
                Me.oTime_Mensajes.Enabled = True
                Call Sleep(3)
                igCnt_CR = igCnt_CR + sgKb_BonC
            End If
        End If
        Me.Refresh_Creditos()
        If Me.otCodigo.Mask <> "##-##-##" Then
            Me.otCodigo.Mask = "##-##-##"
            Me.otCodigo.Tag = ""
            Me.oSetFocus_Codigo.Focus()
        End If
        Exit Sub
    End Sub

    Private Sub Change_Page_Up()
        If Me.oGrpBox_Gen.Visible = True Then
            If (igAct_PgG > 1) Then
                igAct_PgG = igAct_PgG - 1
                Me.Cargar_Pag_Gen(igAct_PgG)
            End If
        End If
        If Me.oGrpBox_Dis.Visible = True Then
            If (igAct_PgD > 1) Then
                igAct_PgD = igAct_PgD - 1
                Me.Cargar_Pag_Dis(igGen_Sel, igAct_PgD)
            End If
        End If

        If Me.oGrpBox_Can.Visible = True Then
            If (igAct_PgC > 1) Then
                igAct_PgC = igAct_PgC - 1
                Me.Cargar_Pag_Canc(igDis_Sel, igAct_PgC)
            End If
        End If
    End Sub

    Private Sub Change_Page_Dn()
        If (Me.oGrpBox_Gen.Visible = True) Then
            If (igAct_PgG <= igMax_Gen) Then
                If igAct_PgG < igMax_RgG Then
                    igAct_PgG = igAct_PgG + 1
                    Me.Cargar_Pag_Gen(igAct_PgG)
                End If
            End If
        End If

        If (Me.oGrpBox_Dis.Visible = True) Then
            If (igAct_PgD <= igMax_Dis) Then
                If igAct_PgD < igMax_RgD Then
                    igAct_PgD = igAct_PgD + 1
                    Me.Cargar_Pag_Dis(igGen_Sel, igAct_PgD)
                End If
            End If
        End If

        If (Me.oGrpBox_Can.Visible = True) Then
            If (igAct_PgC <= igMax_Can) Then
                If igAct_PgC < igMax_RgC Then
                    igAct_PgC = igAct_PgC + 1
                    Me.Cargar_Pag_Canc(igDis_Sel, igAct_PgC)
                End If
            End If
        End If
    End Sub

    Private Sub Desplazar_Pantalla_Gen(ByVal pPage_Order As Char)
        If igMax_RgG = 0 Then
            MsgBox("No hay registros que procesar en esta pàgina", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        Call Desactiva_Genero(True)

        If (UCase(pPage_Order) = sgKb_UP) Then
            If Me.oGrpBox_Gen.Visible = True Then
                If (igAct_PgG > 1) Then
                    igAct_PgG = igAct_PgG - 1
                    Me.Cargar_Pag_Gen(igAct_PgG)
                End If
            End If

        End If

        If (UCase(pPage_Order) = sgKb_DN) Then
            If (Me.oGrpBox_Gen.Visible = True) Then
                If (igAct_PgG <= igMax_Gen) Then
                    If igAct_PgG < igMax_RgG Then
                        igAct_PgG = igAct_PgG + 1
                        Me.Cargar_Pag_Gen(igAct_PgG)
                    End If
                End If
            End If

        End If
        Me.Desactiva_Genero(False)
    End Sub

    Private Sub Desplazar_Pantalla_Dis(ByVal pPage_Order As Char)
        If igMax_RgD = 0 Then
            MsgBox("No hay registros que procesar en esta pàgina", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        Call Desactiva_Disco(True)

        If (UCase(pPage_Order) = sgKb_UP) Then
            If Me.oGrpBox_Dis.Visible = True Then
                If (igAct_PgD > 1) Then
                    igAct_PgD = igAct_PgD - 1
                    Me.Cargar_Pag_Dis(igGen_Sel, igAct_PgD)
                End If
            End If

        End If

        If (UCase(pPage_Order) = sgKb_DN) Then
            If (Me.oGrpBox_Dis.Visible = True) Then
                If (igAct_PgD <= igMax_Dis) Then
                    If igAct_PgD < igMax_RgD Then
                        igAct_PgD = igAct_PgD + 1
                        Me.Cargar_Pag_Dis(igGen_Sel, igAct_PgD)
                    End If
                End If
            End If
        End If
        Me.Desactiva_Disco(False)
    End Sub

    Private Sub Desplazar_Pantalla_Can(ByVal pPage_Order As Char)
        If igMax_RgC = 0 Then
            'MsgBox "No hay canciones en esta página", vbInformation + vbOKOnly
            Exit Sub
        End If

        If (UCase(pPage_Order) = sgKb_UP) Then
            If Me.oGrpBox_Can.Visible = True Then
                If (igAct_PgC > 1) Then
                    igAct_PgC = igAct_PgC - 1
                    Me.Cargar_Pag_Canc(igDis_Sel, igAct_PgC)
                End If
            End If
        End If

        If (UCase(pPage_Order) = sgKb_DN) Then
            If (Me.oGrpBox_Can.Visible = True) Then
                If (igAct_PgC <= igMax_Can) Then
                    If igAct_PgC < igMax_RgC Then
                        igAct_PgC = igAct_PgC + 1
                        Me.Cargar_Pag_Canc(igDis_Sel, igAct_PgC)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Salvar_Temas()
        Dim i As Integer = 0
        Dim cFile As String = Application.StartupPath & "\Reproduccion.txt"
        If IO.File.Exists(cFile) = True Then
            My.Computer.FileSystem.DeleteFile(cFile)
        End If

        Dim s As New IO.StreamWriter(cFile, False)
        For i = 0 To Me.oLst_A_Tocar.Items.Count - 1
            s.WriteLine(Me.oLst_A_Tocar.Items.Item(i))
        Next
        s.Close()

        oIni.SetKeyValue("GENERAL", "IDX_PUBLIC", Format(igInd_Pub, "#####0"))
        oIni.SetKeyValue("CREDITS", "CEDIT_ACAC", Scramble(Format(igCnt_CRG, "#####0")))
    End Sub

    Private Sub Cargar_Temas()
        Dim cFile As String = Application.StartupPath & "\Reproduccion.txt"

        If IO.File.Exists(cFile) = True Then
            Dim lines = File.ReadAllLines(cFile)
            Me.oLst_A_Tocar.Items.Clear()
            Me.oLst_A_Tocar.Items.AddRange(lines)
        End If
    End Sub

    Private Sub otCargador_Music_Tick(sender As Object, e As EventArgs) Handles otCargador_Music.Tick
        On Error GoTo Solve_error
        Dim iCont As Integer
        Dim pCa_Ato As String, sCadenas As String
        If Me.oLst_A_Tocar.Items.Count = 0 Then
            Me.oImg_c_Video.Visible = False
            If igDelay_Bonus_Vid > 0 Then
                If igNext_Bonus <= (Hour(Now()) * 60) + Minute(Now()) Then
                    If (oFlsPub1.Count <= 0) Then
                        igInd_Bon = igInd_Bon + 1
                        Me.oLst_Temas_Video.SelectedIndex = igInd_Bon - 1
                        sCadenas = "999999,999999,999999, 999999, VIDEO BONUS," & sgDir_Pub1 & "\" & oFlsPub1.GetType.FullName.ToString() & ",.,."
                    Else
                        sCadenas = Me.Add_Promos()
                    End If
                    If Not (Strings.Trim(sCadenas) = "") Then
                        Me.oLst_A_Tocar.Items.Add(sCadenas)
                        igNext_Bonus = (Hour(Now()) * 60) + (Minute(Now()) + igDelay_Bonus_Vid)
                    End If
                End If
            End If

            If IO.File.Exists(sgDir_Img & "\Logo1.bmp") Then
                Me.oImg_Logo1.Image = Nothing
                Me.oImg_Logo1.ImageLocation = sgDir_Img & "\Logo1.bmp"
                Me.oImg_Logo1.Visible = True
                Me.MediaPlayer2.close()
                Me.MediaPlayer2.Visible = False

            Else
                Me.oImg_Logo1.Image = Nothing
                Me.oImg_Logo1.ImageLocation = ""
                Me.oImg_Logo1.Visible = False
                Me.MediaPlayer2.Visible = True
            End If
            If igScr_Alone = 0 Then
                If IO.File.Exists(sgDir_Img & "\Logo1.bmp") Then
                    Video_Form.oImg_Logo1.Image = Nothing
                    Video_Form.oImg_Logo1.ImageLocation = sgDir_Img & "\Logo1.bmp"
                    Video_Form.oImg_Logo1.Visible = True
                    Video_Form.MediaPlayer3.close()
                    Video_Form.MediaPlayer3.Visible = False
                Else
                    Video_Form.oImg_Logo1.Image = Nothing
                    Video_Form.oImg_Logo1.ImageLocation = ""
                    Video_Form.MediaPlayer3.Visible = True
                End If
            End If
            'Me.otCodigo.SetFocus
            Exit Sub
        Else
            Me.oImg_Logo1.Visible = False
            Me.oImg_Logo1.Image = Nothing
            Me.oImg_Logo1.ImageLocation = ""
            Me.MediaPlayer2.Visible = True

            If igScr_Alone = 0 Then
                Video_Form.oImg_Logo1.Visible = False
                Video_Form.oImg_Logo1.Image = Nothing
                Video_Form.oImg_Logo1.ImageLocation = ""
                Video_Form.MediaPlayer3.Visible = True
            End If
            Me.Cargar_Musica
            'Me.otCodigo.SetFocus
        End If
        Exit Sub

Solve_error:
        Call ControlError()
        Resume Next

    End Sub

    Public Function SqliteExecute(sql As String) As Boolean
        If Not (oConnSqLite.State = ConnectionState.Open) Then
            oConnSqLite.Open()
        End If

        oCmmSqLite = New SQLiteCommand(sql, oConnSqLite)
        If (oCmmSqLite.ExecuteNonQuery() > 1) Then
            SqliteExecute = True
        Else
            SqliteExecute = False
        End If
    End Function

    Private Function Conectar_DBPromo(Optional bPopulate_ListPub As Boolean = False) As Boolean
        If (bgLoad_Promo = False) Then

            Dim iNum_reg As Integer = 0
            '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------
            Dim sql As String = "" &
            "SELECT " &
            "   file01.ID_GEN , file01.ID_ORD As ID_ORD1, " &
            "   file02.ID_DIS , file02.ID_ORD As ID_ORD2, " &
            "   file02.FL_IMG , file03.ID_CAN, " &
            "   file03.DE_CAN , file03.FL_mp3, file02.FL_IMG   " &
            "FROM file03 " &
            "LEFT JOIN file02 ON file03.id_dis = file02.id_dis " &
            "LEFT JOIN file01 ON file02.id_gen = file02.id_gen " &
            "WHERE((file03.FL_PRC = 1) Or (file02.fl_prd = 1)) " &
            "ORDER by file02.fl_prd, file03.FL_PRC, RANDOM() "

            If Not (oConnSqLite.State = ConnectionState.Open) Then
                oConnSqLite.Open()
            End If

            oCmmSqLite = New SQLiteCommand(sql, oConnSqLite)
            oTableSqLite = oCmmSqLite.ExecuteReader()
            '----------------------------------------------------------------------------------

            If ((oTableSqLite.HasRows = True) And (oTableSqLite.EndOfData = False)) Then
                If bPopulate_ListPub = True Then
                    Me.oLst_Temas_Prom.Items.Clear()
                    While oTableSqLite.Read()
                        Me.oLst_Temas_Prom.Items.Add(oTableSqLite.GetValue(0))
                    End While
                Else
                    oTableSqLite.Read()
                End If
                bgLoad_Promo = True
            Else
                bgLoad_Promo = False
            End If
        Else
            bgLoad_Promo = True
        End If
        Return bgLoad_Promo
    End Function

    Private Sub Conectar_DBPub(Optional bPopulate_ListPub As Boolean = True)
        If (bgLoad_pub = False) Then
            Try
                oDirPub1 = New IO.DirectoryInfo(sgDir_Pub1)
                oFlsPub1 = oDirPub1.GetFiles("*.*")

                If (bPopulate_ListPub = True) Then
                    Dim dra1 As IO.FileInfo
                    Me.oLst_Temas_Pub.Items.Clear()
                    For Each dra1 In oFlsPub1
                        Me.oLst_Temas_Pub.Items.Add(dra1.GetType.FullName.ToString())
                    Next
                    dra1 = Nothing
                    Call Shuffle_ListBox(Me.oLst_Temas_Pub)
                End If

                oDirPub2 = New IO.DirectoryInfo(sgDir_Pub2)
                oFlsPub2 = oDirPub2.GetFiles("*.*")

                If (bPopulate_ListPub = True) Then
                    Dim dra2 As IO.FileInfo
                    Me.oLst_Temas_Pub.Items.Clear()
                    For Each dra2 In oFlsPub2
                        Me.oLst_Temas_Pub.Items.Add(dra2.GetType.FullName.ToString())
                    Next
                    dra2 = Nothing
                    Call Shuffle_ListBox(Me.oLst_Temas_Pub)
                End If

            Catch ex As Exception
                Console.WriteLine("Error: " & ex.ToString())
            End Try
        End If
    End Sub

    Private Function Add_Promos() As String
        Me.Conectar_DBPromo()

        Dim cOrd_Gen As String, cOrd_Dis As String, cOrd_Can As String
        Dim cCo_Can As String, cFl_MP3 As String, cDes_Can As String
        Dim cFl_Dis As String
        Dim cValue As String = ""

        If ((oTableSqLite.HasRows = True) And (oTableSqLite.EndOfData = False)) Then
            oTableSqLite.Read()
            cOrd_Gen = oTableSqLite.Item("ID_ORD1").Value
            cOrd_Dis = oTableSqLite.Item("ID_ORD2").Value
            cOrd_Can = oTableSqLite.Item("ID_ORD").Value
            cCo_Can = oTableSqLite.Item("ID_CAN").Value
            cFl_MP3 = oTableSqLite.Item("FL_mp3").Value
            cFl_Dis = oTableSqLite.Item("FL_IMG").Value
            cDes_Can = oTableSqLite.Item("DE_CAN").Value
            cValue = Strings.Trim(cOrd_Gen & cOrd_Dis & cOrd_Can) & "," & Strings.Trim(cCo_Can) & "," & Strings.Trim(cDes_Can) & "," & Strings.Trim(cFl_MP3) & "," & "*" & "," & Strings.Trim(cFl_Dis)
            Return cValue
        Else
            Return ""
        End If
    End Function

    Private Sub Cargar_Musica()
        Dim sFle_Mp3 As String
        Dim aDet() As String
        If Me.oLst_A_Tocar.Items(0) <> "" Then
            aDet = Strings.Split(Me.oLst_A_Tocar.Items(0), ",", , vbTextCompare)
            sFle_Mp3 = Strings.Trim(Strings.Trim(aDet(3)))
            If (UCase(Strings.Right(Strings.Trim(sFle_Mp3), 3)) <> "MP3") Then
                bgIs_Video = True
                Me.oImg_c_Video.Visible = True
            Else
                bgIs_Video = False
                Me.oImg_c_Video.Visible = False
            End If
            If IO.File.Exists(sFle_Mp3) Then
                Me.Cargar_Musica_P0(sFle_Mp3)
            Else
                If (sFle_Mp3 <> "") Then
                    Me.olMessage.Text = "Tema no encontrado!"
                    Me.olMessage.Visible = True
                    Me.oTime_Mensajes.Enabled = True
                    Me.oImg_c_Video.Visible = False
                    Call Retrocede()
                End If
            End If
        Else
            oImg_c_Video.Visible = False
        End If
        Me.Muestra_Tema_Det()
    End Sub

    Private Sub Cargar_Musica_P0(ByVal spFle_Mp3 As String)
        If bgIs_Video = True Then
            If (Video_Form.MediaPlayer3.URL <> spFle_Mp3) Then
                Me.MediaPlayer2.close()
            Else
                If ObPlayer_Ocupado(Video_Form.MediaPlayer3) = True Then
                    Exit Sub
                End If
            End If
            '********************VISOR DE VIDEO GRANDE****************************
            Video_Form.MediaPlayer3.URL = spFle_Mp3
            Video_Form.MediaPlayer3.settings.mute = False
            Video_Form.MediaPlayer3.settings.volume = 120
            Video_Form.MediaPlayer3.Ctlcontrols.play()
            '********************TOCADOR DE MUSICA SOLA***************************
            Me.MediaPlayer1.URL = ""
            Me.MediaPlayer1.settings.mute = True
            '*********************VISOR DE VIDEO CHICO****************************
            Me.MediaPlayer2.URL = ""
            Me.MediaPlayer2.settings.mute = True
        Else
            If ObPlayer_Ocupado(Me.MediaPlayer1) = True Then
                Exit Sub
            End If
            '********************VISOR DE VIDEO GRANDE****************************
            Video_Form.MediaPlayer3.URL = ""
            Video_Form.MediaPlayer3.settings.mute = True
            '********************TOCADOR DE MUSICA SOLA***************************
            Me.MediaPlayer1.URL = spFle_Mp3
            Me.MediaPlayer1.settings.mute = False
            Me.MediaPlayer1.settings.volume = 120
            '*********************VISOR DE VIDEO CHICO****************************
            Me.MediaPlayer2.URL = ""
            Me.MediaPlayer2.settings.mute = True
            '***********************************************
        End If
    End Sub

    Public Function ObPlayer_Ocupado(oPlayer As AxWMPLib.AxWindowsMediaPlayer) As Boolean
        If (oPlayer.playState = WMPLib.WMPPlayState.wmppsWaiting Or
            oPlayer.playState = WMPLib.WMPPlayState.wmppsPlaying Or
            oPlayer.playState = WMPLib.WMPPlayState.wmppsBuffering Or
            oPlayer.playState = WMPLib.WMPPlayState.wmppsTransitioning) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If igLen > 0 Then
            If (Val(Me.olTimer1.Text) > 0) Then
                Me.olTimer1.Text = PADL(Conversion.Str(Conversion.Val(olTimer1.Text) - 1).Trim, 3, "0")
            Else
                Me.olTimer1.Text = 0
                Me.olTimer1.Visible = False
            End If

            If (Val(Me.olTimer2.Text) > 0) Then
                Me.olTimer2.Text = PADL(Conversion.Str(Conversion.Val(olTimer2.Text) - 1).Trim, 3, "0")
            Else
                Me.olTimer2.Text = 0
                Me.olTimer2.Visible = False
            End If
        End If
        Application.DoEvents()
    End Sub

    Private Sub oTimer_Srv_Tick(sender As Object, e As EventArgs) Handles oTimer_Srv.Tick
        If (Me.olMessage.Tag = "1") Then
            Me.oTime_Mensajes.Tag = ""
        End If
        Me.oTimer_Srv.Enabled = False
    End Sub

    Private Sub Muestra_Tema_Det()
        Dim aDet() As String
        If (Me.oLst_A_Tocar.Items.Count = 0) Then
            Me.otTema_Act.Text = ""
            Me.olTema_Act.Text = ""
            Exit Sub
        End If
        aDet = Strings.Split(Me.oLst_A_Tocar.Items(0), ",", , vbTextCompare)
        Me.otTema_Act.Text = aDet(0)
        If (Strings.Trim(aDet(4)) = "*") Then
            Me.olCred_Msg.Tag = Me.olCred_Msg.Text
            Me.olCred_Msg.Text = "DISCO PROMO..."
            Me.Image2.ImageLocation = aDet(5)
        Else
            If Me.olCred_Msg.Text <> "INSERTE ¢ 0.25" Then
                Me.olCred_Msg.Text = "INSERTE ¢ 0.25"
            End If
            Me.Refresh_Creditos()
        End If
        Me.olTema_Act.Text = Strings.Trim(aDet(2))
    End Sub

    Private Sub Remove_Temes()
        If (Me.oLst_A_Tocar.Items.Count > 0) Then
            Me.oLst_A_Tocar.Items.RemoveAt(0)
            Me.oLst_A_Tocar.Refresh()
            Me.Muestra_Tema_Det()
        End If
    End Sub

    Private Sub oTM_Box_Tick(sender As Object, e As EventArgs) Handles oTM_Box.Tick

    End Sub
End Class

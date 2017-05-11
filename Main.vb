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
        REM Call Me.Limpia_Gen()
        REM Call Me.Limpia_Dis()
        REM Call Me.Limpia_Can()
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
        Dim pFileName = Application.StartupPath() & "\FilesV3.db;FailIfMissing=False;"
        Dim str As String = "DataSource=" & pFileName
        Dim con As New SQLiteConnection(str)
        Dim sql As String = "SELECT id_gen, id_ord, descri FROM file01 WHERE page=" & pNum_Pag & " ORDER BY id_ord"
        Dim dar As SQLiteDataReader

        If Not (con.State = ConnectionState.Open) Then
            con.Open()
        End If

        Dim cmm As New SQLiteCommand(sql, con)
        dar = cmm.ExecuteReader()
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
        con.Close()
        con = Nothing
        cmm = Nothing
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
        Dim pFileName = Application.StartupPath() & "\FilesV3.db;FailIfMissing=False;"
        Dim str As String = "DataSource=" & pFileName
        Dim con As New SQLiteConnection(str)
        Dim sql As String = "SELECT * FROM file02 WHERE id_gen=" & pCod_Gen & " AND page=" & pNum_Pag & " ORDER BY id_gen, id_ord"
        Dim dar As SQLiteDataReader

        If Not (con.State = ConnectionState.Open) Then
            con.Open()
        End If

        Dim cmm As New SQLiteCommand(sql, con)
        dar = cmm.ExecuteReader()
        '----------------------------------------------------------------------------------

        Me.oLTitulo.Text = "[SELECCIONE EL DISCO.]"

        '------------------------MOSTRAR DATOS EN PANTALLA---------------------------------
        Dim sFile As String = ""
        Dim sLabel As String = ""
        Dim i As Integer = 0
        Dim cPahFile As String = ""

        If dar.HasRows = True Then
            iNum_reg = 1
            Do While dar.Read()

                sFile = Trim(System.IO.Path.GetFileName(Trim(dar.Item("FL_IMG").ToString)))
                cPahFile = sParam(4) + "\" + Trim(System.IO.Path.GetFileName(Trim(dar.Item("FL_IMG").ToString)))
                sLabel = Trim(dar.Item("ID_ORD").ToString)
                If Microsoft.VisualBasic.Right(cPahFile, 1) = "\" Then
                    cPahFile = ""
                End If

                Me.Disc_Num(iNum_reg).TAG = Val(dar.Item("id_dis").ToString())
                Me.Disc_Num(iNum_reg).text = sLabel
                Me.Disc_Num(iNum_reg).cursor = Cursors.Hand
                Me.Disc_Num(iNum_reg).Visible = True

                Me.Disc_Img(iNum_reg).TAG = Val(dar.Item("id_dis").ToString())
                Me.Disc_Img(iNum_reg).cursor = Cursors.Hand

                If (Trim(sFile) = "") Then
                    Me.Disc_Img(iNum_reg).ImageLocation = ""
                    Me.Disc_Img(iNum_reg).Image = My.Resources.Resources.sin_images
                Else
                    If My.Computer.FileSystem.FileExists(cPahFile) = True Then
                        Me.Disc_Img(iNum_reg).Image = Nothing
                        Me.Disc_Img(iNum_reg).ImageLocation = cPahFile
                    Else
                        Me.Disc_Img(iNum_reg).ImageLocation = ""
                        Me.Disc_Img(iNum_reg).Image = My.Resources.Resources.no_images
                    End If
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
                    Me.oDisc_LabelA(iNum_reg).text = UCase(Trim(PADL(dar.Item("NOM_DIS").ToString, 19, " ")))
                    Me.oDisc_LabelB(iNum_reg).Visible = True
                    Me.oDisc_LabelB(iNum_reg).text = UCase(Trim(PADL(dar.Item("NOM_ART").ToString, 19, " ")))
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
        con.Close()
        con = Nothing
        cmm = Nothing
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
        Dim pFileName = Application.StartupPath() & "\FilesV3.db;FailIfMissing=False;"
        Dim str As String = "DataSource=" & pFileName
        Dim con As New SQLiteConnection(str)
        Dim sql As String = "SELECT * FROM file03 WHERE id_dis=" & pCod_Dis & " AND page=" & pNum_Pag & " ORDER BY id_dis, id_ord"
        Dim dar As SQLiteDataReader

        If Not (con.State = ConnectionState.Open) Then
            con.Open()
        End If

        Dim cmm As New SQLiteCommand(sql, con)
        dar = cmm.ExecuteReader()
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
                sNamFl = System.IO.Path.GetFileName(Trim(dar.Item("fl_mp3").ToString))
                sExtFl = System.IO.Path.GetExtension(Trim(dar.Item("fl_mp3").ToString))

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

                If My.Computer.FileSystem.FileExists(sFile) = True Then
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

    Private Function Busca_Sel_1(ByVal pValor_Bus As String) As Boolean
        Dim i As Integer
        Busca_Sel_1 = False
        For i = 1 To Me.Gene_Lst.Count
            If Me.Gene_Lst(i).tag = pValor_Bus Then
                igGen_Sel = Me.Gene_Lst(i).tag
                Busca_Sel_1 = True
                Exit Function
            End If
        Next
    End Function


    Private Sub AxVLCPlugin21_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        End
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)

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
        igMax_Gen = 13 - 2 : igMax_Dis = 13 : igMax_Can = 13 - 2 'valores fijos su valor no pueded ser superior
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
        Me.oLVersion.Text = "Ver." & Trim(Application.ProductVersion)

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
        sgNom_Loc = Trim(sParam(10))
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
        '    sVal1 = Scramble(Trim(ReadText()))
        '    sVal2 = Trim(oIni.GetKeyValue("SERIAL", "ID", ""))
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
        Call Refresh_Creditos(Me)
        '----------------------GENEROS------------------------------

        Me.Cargar_Gen()
        Me.Desactiva_Genero(False)
        '---------------------PUBLICIDAD----------------------------
        REM Call Conectar_DBPub()
        '-----------------------PROMOS------------------------------
        REM Call Conectar_DBPro()
        '-----------------------Otros-------------------------------
        'Me.otNot_Found_List.Text = ""


        'Call Cargar_Temas
        'If igScr_Alone = 0 Then
        'Video_Form.Show
        'End If
        'If bgKeep_On_Top = True Then
        ' Call AlwaysOnTop(Main_Form, True)
        'End If
        'If igStartPlayMusic = 1 Then
        'My.Computer.Keyboard.SendKeys(sgKb_ResA)
        'Call Salvar_Temas
        'End If
        Exit Sub

Solve_error:
        Call ControlError()
        Resume Next
    End Sub

    Private Sub Check_Other()
        If Not sgFle_Fon = "" Then
            If My.Computer.FileSystem.FileExists(sgFle_Fon) Then
                Me.BackgroundImage = System.Drawing.Image.FromFile(sgFle_Fon)
                'aqui
            End If
        End If
        On Error Resume Next
        Call MkDir(sgDir_Tmp)
        On Error Resume Next
        Call MkDir(sgDir_odb)
        '---------------------------------------Link_Tab.dsn---------------------------------------
        Dim oIni2 As New IniFile

        oIni2.Load(sgDir_odb & "\Link_Tab.dsn")
        If Not My.Computer.FileSystem.FileExists(sgDir_odb & "\Link_Tab.dsn") Then
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
        If Not My.Computer.FileSystem.FileExists(sgDir_odb & "\Link_Dbf.dsn") Then
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
        If Not My.Computer.FileSystem.FileExists(sgDir_Fls & "\schema.ini") Then
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

            If Not My.Computer.FileSystem.FileExists(sgDir_Fls & "\file05.tab") Then
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
        If Me.oGrpBox_Gen.Visible = True Then
            If (igAct_PgG > 1) Then
                igAct_PgG = igAct_PgG - 1
                Call Cargar_Pag_Gen(igAct_PgG)
            End If
        End If
        If Me.oGrpBox_Dis.Visible = True Then
            If (igAct_PgD > 1) Then
                igAct_PgD = igAct_PgD - 1
                Call Cargar_Pag_Dis(igGen_Sel, igAct_PgD)
            End If
        End If

        If Me.oGrpBox_Can.Visible = True Then
            If (igAct_PgC > 1) Then
                igAct_PgC = igAct_PgC - 1
                Call Cargar_Pag_Canc(igDis_Sel, igAct_PgC)
            End If
        End If
    End Sub

    Private Sub Img_PagNext_Click(sender As Object, e As EventArgs) Handles Img_PagNext.Click
        If (Me.oGrpBox_Gen.Visible = True) Then
            If (igAct_PgG <= igMax_Gen) Then
                If igAct_PgG < igMax_RgG Then
                    igAct_PgG = igAct_PgG + 1
                    Call Cargar_Pag_Gen(igAct_PgG)
                End If
            End If
        End If

        If (Me.oGrpBox_Dis.Visible = True) Then
            If (igAct_PgD <= igMax_Dis) Then
                If igAct_PgD < igMax_RgD Then
                    igAct_PgD = igAct_PgD + 1
                    Call Cargar_Pag_Dis(igGen_Sel, igAct_PgD)
                End If
            End If
        End If

        If (Me.oGrpBox_Can.Visible = True) Then
            If (igAct_PgC <= igMax_Can) Then
                If igAct_PgC < igMax_RgC Then
                    igAct_PgC = igAct_PgC + 1
                    Call Cargar_Pag_Canc(igDis_Sel, igAct_PgC)
                End If
            End If
        End If
    End Sub

    Private Sub oTime_Mensajes_Tick(sender As Object, e As EventArgs) Handles oTime_Mensajes.Tick
        If Me.oTimer_Reset.Enabled = False Then
            Me.oTimer_Reset.Enabled = True
        End If
        If (Me.olMessage.ForeColor) = ConvertToRbg(&HFF&) Then
            Me.olMessage.ForeColor = ConvertToRbg(&HFFFF&)
        Else
            Me.olMessage.ForeColor = ConvertToRbg(&HFF&)
        End If
        Application.DoEvents()
    End Sub

    Private Sub oTimer_Reset_Tick(sender As Object, e As EventArgs) Handles oTimer_Reset.Tick
        Call Muestra_Tema_Det()
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
        If ((Trim(Me.otCodigo.Text) = "") Or (igNext_Return_Gen = 0)) Then
            Exit Sub
        End If
        If igDelay_Return_Gen > 0 Then
            'If igCan_Sel = "" Then
            'Sólo regresa a género si esta en la pantalla de discos
            If igNext_Return_Gen <= (Hour(Now()) * 60) + Minute(Now()) Then
                Call Retrocede
                Call Retrocede
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
                'My.Computer.Keyboard.SendKeys("{BACKSPACE}")
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

        Application.DoEvents()
    End Sub

    Private Sub otCodigo_TextChanged(sender As Object, e As EventArgs) Handles otCodigo.TextChanged
        REM On Error GoTo Solve_error
        Dim sValue As String = Me.otCodigo.Text.Trim().Replace("-", "")
        igLen = sValue.Trim.Length


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
                Call Cargar_Gen()
            Case 1 To 2
                If gbServ_Mode = True Then
                    Call Desactiva_Cancion(True)
                    Call Desactiva_Disco(True)
                    Call Desactiva_Genero(True)
                    Me.oService_Info.LoadFile(Application.StartupPath & "\Service_Info.RTF")
                    Me.olT_Mant.Visible = True
                    Me.oService_Info.Visible = True
                    Exit Sub
                End If
                Dim sGen_Ret As String
                sGen_Ret = ""
                igGen_Sel = ""
                igDis_Sel = ""
                igCan_Sel = ""

                Call Desactiva_Cancion(True)
                Call Desactiva_Disco(True)
                Call Desactiva_Genero(False)
                Me.Image2.ImageLocation = ""
                If igLen < 2 Then
                    Exit Sub
                End If
                If Busca_Sel_1(sValue) = False Then
                    If igLen = 2 Then
                        Me.oLTitulo.Text = "[GENERO NO EXISTE!!!]"
                    Else
                        Me.oLTitulo.Text = ""
                    End If
                    Call Retrocede()
                    igAct_PgG = 1
                    Exit Sub
                Else
                    Me.oLTitulo.Text = "[Seleccione le Disco]..."
                End If
                igAct_PgD = 1
                igAct_PgC = 1
                '***********************DISCOS************************
                Call Cargar_Pag_Dis(igGen_Sel, igAct_PgD)
                bgVIP = False
                igNext_Return_Gen = (Hour(Now()) * 60) + (Minute(Now()) + igDelay_Return_Gen)
                If Val(Me.olTimer1.Text) = 0 Then
                    Me.olTimer1.Text = Format(igDelay_Return_Gen * 60, "#0")
                    Me.olTimer1.Visible = True
                End If

        End Select
    End Sub

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

            Call Retrocede()
            Call Retrocede()
        End If
    End Sub
End Class

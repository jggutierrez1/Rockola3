Option Explicit On
Imports System
Imports System.Collections
Imports System.Management
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Threading
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports Devart.Data.SQLite
Imports System.Drawing

Module Module1

    Public Structure POINTAPI
        Public X As Long
        Public Y As Long
    End Structure

    Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

    Public oIni As New IniFile
    Private bini_started As Boolean = False
    Public cini_path As String = ""

    Public igFlg_TCR As Integer
    Public igShowPass As Integer
    Public bgSw As Boolean
    Public bgBlinkPag As Boolean
    Public agArr_Pub() As String
    Public sgParms() As String
    Public bgVideoLabel As Boolean
    Public bgDiscLabel As Boolean
    Public sgCmdLine As String
    Public bgExit As Boolean
    Public igInd_Kar As Integer
    Public bgPopular As Boolean  'Si esta activado al orden de popular
    Public bgVIP As Boolean      'Si esta activado al orden de VIP
    Public igKeyAscii As Integer 'Almacena el còdigo de la ùltima letra presionada

    Public igMax_Dis As Integer = 12 'Máxima cantidad de registros (DB) físicos en Discos
    Public igMax_Gen As Integer = 13 'Máxima cantidad de registros (DB) físicos en Géneros

    Public igMax_Can As Integer = 13 'Máxima cantidad de registros (DB) físicos en Géneros
    Public igMax_RgG As Integer  'Máximo de registros por pantalla de Género

    Public igMax_RgD As Integer  'Máximo de registros por pantalla de Discos
    Public igMax_RgC As Integer  'Máximo de registros por pantalla de Discos

    Public igTot_PgG As Integer  'Total de paginas en la consulta (Generos)
    Public igAct_PgG As Integer  'Página actual (Generos)

    Public igTot_PgD As Integer  'Total de paginas en la consulta (Discos)
    Public igAct_PgD As Integer  'Página actual (Discos)

    Public igTot_PgC As Integer  'Total de paginas en la consulta (Discos)
    Public igAct_PgC As Integer  'Página actual (Cansión)

    Public sgIdx_Prm As Integer  'Máximo de temas para insertar un promo
    Public igFlg_SavedCR As Integer
    Public igStartPlayMusic As Integer
    Public sgCr_AKey As String
    Public igKeep_Cred As Integer
    Public igNoDuplicT As Integer
    Public igMixe_Popu As Integer
    Public igLeftDisk As Integer
    Public sgKb_Crd1 As String
    Public sgKb_Crd2 As String
    Public sgKb_ResM As String
    Public sgKb_ResA As String
    Public sgKb_BonC As String
    Public sgKb_VID As String
    Public sgKb_Del As String
    Public sgKb_Ret As String
    Public sgKb_Pop As String
    Public sgKb_VIP As String
    Public sgKb_Vef As String
    Public sgKb_UP As String
    Public sgKb_Pause As String
    Public sgKb_SwK As String
    Public sgKb_DN As String
    Public sgKb_SwP As String
    Public sgFec_iAc As String
    Public sgFec_Fac As String
    Public sgSer_Mac As String
    Public sgSer_CPU As String
    Public sgNom_Loc As String
    Public sgWin_Key As String
    Public bFlagFoc As Boolean
    Public sgDir_odb As String
    Public sgDir_Tmp As String
    Public sgDir_Fls As String
    Public sgDir_Fls2 As String
    Public sgDir_Img As String
    Public sgDir_Mp3 As String
    Public sgDir_Pub1 As String    'Directorio de publicidad #1
    Public sgDir_Pub2 As String    'Directorio de publicidad #2
    Public sgFle_Fon As String
    Public igLim_Cred As Integer    'Limite de creditos
    Public igInd_Pub As Integer     'índice de publicidad

    Public sgGen_Sel_Ord As String
    Public sgDis_Sel_Ord As String
    Public sgCan_Sel_Ord As String

    Public sgGen_Sel_Id As String
    Public sgDis_Sel_Id As String
    Public sgCan_Sel_Id As String

    Public igCnt_CR As Integer      'Contador de créditos.
    Public igCnt_CRP As Integer     'Contador de créditos de prueba
    Public igCnt_CRG As Long        'Contador de Créditos general.
    Public sParam(0 To 53) As String

    Public igNext_Return_Gen As Integer
    Public igDelay_Return_Gen As Integer
    Public igDelay_Return_Dis As Integer
    Public igNext_Bonus As Integer
    Public igLen As Integer
    Public bgSw_Pub As Boolean
    Public igInd_Bon As Integer     'Indicador de número de bonus
    Public igTot_Pub As Integer
    Public vgForeColor
    Public vgForeColorI
    Public vTemp
    Public igDelay_Del_Dig_Can As Integer
    Public igDelay_Ret_Gen As Integer
    Public igNo_RgAt As Integer

    Public bgKeep_On_Top As Boolean
    Public bgIs_Video As Boolean
    Public bgIs_Publi As Boolean
    Public bgWMP_Busy As Boolean
    Public igCont_Sin As Integer
    Public igDelay_Bonus_Vid As Integer
    Public rtn As Long
    Public igScr_Alone As Integer

    Public bgLoad_Promo As Boolean = False
    Public bgLoad_pub As Boolean = False

    Public oDirPub1 As IO.DirectoryInfo
    Public oFlsPub1 As IO.FileInfo()
    Public oDirPub2 As IO.DirectoryInfo
    Public oFlsPub2 As IO.FileInfo()

    Public cFileSqlite As String = Application.StartupPath() & "\FilesV3.db;FailIfMissing=False;"
    Public cStrConnect As String = "DataSource=" & cFileSqlite
    Public oConnSqLite As New SQLiteConnection(cStrConnect)
    Public oCmmSqLite As SQLiteCommand
    Public oTableSqLite As SQLiteDataReader

    Public iNivel_Pag As Integer = 1


    'Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
    'ByVal nSize As Long) As Long
    Public Const MAX_PATH = 255

    Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

    Public igMainCode As String

    Public Const SWP_HIDEWINDOW = &H80
    Public Const SWP_SHOWWINDOW = &H40
    'Control Panel
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

    '=====================================LLamado a librería para desactivar screen saver==========================

    Public Enum EXECUTION_STATE As UInteger ' Determine Monitor State
        ES_AWAYMODE_REQUIRED = &H40
        ES_CONTINUOUS = &H80000000UI
        ES_DISPLAY_REQUIRED = &H2
        ES_SYSTEM_REQUIRED = &H1
        ' Legacy flag, should not be used.
        ' ES_USER_PRESENT = 0x00000004
    End Enum

    'Enables an application to inform the system that it is in use, thereby preventing the system from entering sleep or turning off the display while the application is running.
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Function SetThreadExecutionState(ByVal esFlags As EXECUTION_STATE) As EXECUTION_STATE
    End Function

    'This function queries or sets system-wide parameters, and updates the user profile during the process.
    <DllImport("user32", EntryPoint:="SystemParametersInfo", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Function SystemParametersInfo(ByVal uAction As Integer, ByVal uParam As Integer, ByVal lpvParam As String, ByVal fuWinIni As Integer) As Integer
    End Function

    Public Const SPI_SETSCREENSAVETIMEOUT As Int32 = 15
    Public Const SPI_GETSCREENSAVEACTIVE As Long = &H10
    Public Const SPI_GETSCREENSAVERRUNNING As Long = &H72
    '================================================================================================================

    Public Function ConvertToRbg(ByVal HexColor As String) As Color
        Dim Red As String
        Dim Green As String
        Dim Blue As String
        HexColor = Replace(HexColor, "#", "")
        Red = Val("&H" & Mid(HexColor, 1, 2))
        Green = Val("&H" & Mid(HexColor, 3, 2))
        Blue = Val("&H" & Mid(HexColor, 5, 2))
        Return Color.FromArgb(Red, Green, Blue)
    End Function

    Public Sub Execute_SqlLite(pFileName As String, cSqlCmd As String)
        If pFileName = "" Then
            pFileName = Application.StartupPath() & "\FilesV3.tab;FailIfMissing=False;"
        End If

        Dim cConnString As String = "DataSource=" & pFileName

        Dim oConn As New SQLiteConnection(cConnString)

        Dim oComm As New SQLiteCommand(cSqlCmd)
        oComm.Connection = oConn
        oConn.Open()
        Try
            Dim aff As Integer = oComm.ExecuteNonQuery()
            Debug.Print(aff & " rows were affected.")
        Catch
            MessageBox.Show("Error encountered during INSERT operation.")
        Finally
            oConn.Close()
        End Try
        oConn = Nothing
        oComm = Nothing
    End Sub

    Sub SqlLite_Verify_Data()
        Dim cSqlStr As String
        cSqlStr = "" &
            "CREATE TABLE IF NOT EXISTS file01( " &
            "id_gen Integer(11) Not NULL," &
            "id_orda VARCHAR(3) Not NULL," &
            "id_ord VARCHAR(3) Not NULL," &
            "descri VARCHAR(20) Not NULL," &
            "gen_st Integer(11) Not NULL," &
            "ult_act DateTime Not NULL," &
            "mark VARCHAR(1) Not NULL) "
        Call Execute_SqlLite("", cSqlStr)

        cSqlStr = "" &
            "CREATE TABLE IF NOT EXISTS file02(" &
            "id_gen Integer(11) Not NULL," &
            "id_dis Integer(11) Not NULL," &
            "id_ord VARCHAR(3) Not NULL," &
            "id_orda VARCHAR(3) Not NULL," &
            "nom_art VARCHAR(40) Not NULL," &
            "nom_dis VARCHAR(40) Not NULL," &
            "fl_img VARCHAR(80) Not NULL," &
            "tx_com VARCHAR(40) Not NULL," &
            "mark VARCHAR(1) Not NULL," &
            "dis_st Integer(11) Not NULL," &
            "mp3_err VARCHAR(1) Not NULL," &
            "img_err VARCHAR(1) Not NULL," &
            "c_video Decimal(2, 0) Not NULL," &
            "ult_act DateTime Not NULL," &
            "fl_new Decimal(2, 0) Not NULL," &
            "fl_prd Decimal(2, 0) Not NULL," &
            "counter Double Not NULL )"
        Call Execute_SqlLite("", cSqlStr)

        cSqlStr = "" &
            "CREATE TABLE IF NOT EXISTS file03(" &
            "id_dis INT(11) Not NULL," &
            "id_can INT(11) Not NULL," &
            "id_ord Char(3) Not NULL," &
            "de_can Char(50) Not NULL," &
            "fl_mp3 Char(80) Not NULL," &
            "mark Char(1) Not NULL," &
            "ult_act DATETIME Not NULL," &
            "fl_prc Decimal(2,0) Not NULL," &
            "counter Double Not NULL," &
            "c_video Char(1) Not NULL," &
            "fle_sec Decimal(16,0) Not NULL)"
        Call Execute_SqlLite("", cSqlStr)

        cSqlStr = "" &
            "CREATE TABLE IF NOT EXISTS file05(" &
            "id_reg INT(11) Not NULL," &
            "id_can INT(11) Not NULL," &
            "id_org CHAR(10) Not NULL," &
            "de_can CHAR(50) Not NULL," &
            "fl_mp3 CHAR(80) Not NULL," &
            "sec1 CHAR(4) Not NULL," &
            "sec2 CHAR(4) Not NULL," &
            "sec3 CHAR(4) Not NULL," &
            "vip CHAR(4) Not NULL," &
            "video CHAR(4) Not NULL," &
            "id_tipo CHAR(4) Not NULL," &
            "id_ord CHAR(10) Not NULL)"
        Call Execute_SqlLite("", cSqlStr)

        'Call Execute_SqlLite("", "TRUNCATE file01")
        'Call Execute_SqlLite("", "TRUNCATE file02")
        'Call Execute_SqlLite("", "TRUNCATE file03")
        'Call Execute_SqlLite("", "TRUNCATE file05")
    End Sub

    Public Sub Init_Ini()
        Dim sRuta As String = ""
        sRuta = Application.StartupPath()
        If (IO.File.Exists(sRuta & "\PathV3.ini") = True) Then
            cini_path = sRuta & "\PathV3.ini"
            oIni.Load(cini_path)
            bini_started = True
        Else
            cini_path = ""
            bini_started = False
            MsgBox("Archivo ini [" & sRuta & "\PathV3.ini" & "] NO EXISTE", 0)
        End If
    End Sub

    Public Sub Save_Defa_Path(Optional ByVal pRuta As String = "")
        If pRuta = "" Then
            pRuta = Application.StartupPath()
        End If
        If (bini_started = True) Then

            oIni.SetKeyValue("PATHS", "DIR_ODB", pRuta & "\ODBC")
            oIni.SetKeyValue("PATHS", "DIR_ODB", pRuta & "\ODBC")
            oIni.SetKeyValue("PATHS", "DIR_TMP", pRuta & "\TMP")
            oIni.SetKeyValue("PATHS", "DIR_FL1", pRuta & "\FILES1")
            oIni.SetKeyValue("PATHS", "DIR_FL2", pRuta & "\FILES2")
            oIni.SetKeyValue("PATHS", "DIR_IMG", pRuta & "\fotos")
            oIni.SetKeyValue("PATHS", "DIR_MP3", pRuta & "\cancionero")
            oIni.SetKeyValue("PATHS", "DIR_PUB1", pRuta & "\PUBLICIDAD1")
            oIni.SetKeyValue("PATHS", "DIR_PUB2", pRuta & "\PUBLICIDAD2")
            '----------------------------------------------------------------------------------'
            oIni.SetKeyValue("TIMES", "TM_RET_GEN", "03")
            oIni.SetKeyValue("TIMES", "TM_RET_DIS", "02")
            oIni.SetKeyValue("TIMES", "TM_BON_VID", "20")
            oIni.SetKeyValue("TIMES", "ID_CNT_PRO", "0")
            '----------------------------------------------------------------------------------'
            oIni.SetKeyValue("ROCKOLA", "FILE_BACKG", pRuta & "\fondo8.JPG")
            oIni.SetKeyValue("ROCKOLA", "SERIAL_MAC", "")
            oIni.SetKeyValue("ROCKOLA", "SERIAL_CPU", "")
            oIni.SetKeyValue("ROCKOLA", "NOMBRE_LOC", Scramble("SIN ASIGNACIÓN!"))
            oIni.SetKeyValue("ROCKOLA", "FECHA_ACTI", "")
            oIni.SetKeyValue("ROCKOLA", "FECHA_ACTF", "")
            oIni.SetKeyValue("ROCKOLA", "WAPLIC_KEY", "")
            oIni.SetKeyValue("ROCKOLA", "CR_ACC_KEY", "")
            oIni.SetKeyValue("ROCKOLA", "SHOW_VIDLA", "0")
            oIni.SetKeyValue("ROCKOLA", "SHOW_DISLA", "0")
            oIni.SetKeyValue("ROCKOLA", "PASSW__BOX", "0")
            oIni.SetKeyValue("ROCKOLA", "RELOAD_APP", "0")
            oIni.SetKeyValue("ROCKOLA", "APPRUNNING", "0")
            oIni.SetKeyValue("ROCKOLA", "START_MLST", "0")
            '----------------------------------------------------------------------------------'
            oIni.SetKeyValue("GENERAL", "SHOW_ONTOP", "0")
            oIni.SetKeyValue("GENERAL", "SCRN_ALONE", "1")
            oIni.SetKeyValue("GENERAL", "NDUP_TEMES", "0")
            oIni.SetKeyValue("GENERAL", "SWITCH_PUB", "1")
            oIni.SetKeyValue("GENERAL", "SWITCH_KAR", "0")
            '----------------------------------------------------------------------------------'
            oIni.SetKeyValue("CREDITS", "ACU_SAVECR", "0")
            oIni.SetKeyValue("CREDITS", "FLG_SAVECR", "1")
            oIni.SetKeyValue("CREDITS", "KEEP_SCRED", "0")
            oIni.SetKeyValue("CREDITS", "LIMCR_CRED", "20")
            oIni.SetKeyValue("CREDITS", "BONUS_CRED", "0")
            oIni.SetKeyValue("CREDITS", "VIDEO_CRED", "0")
            oIni.SetKeyValue("CREDITS", "CEDIT_ACAN", "0")
            oIni.SetKeyValue("CREDITS", "CEDIT_ACAC", "0")
            '----------------------------------------------------------------------------------'
            oIni.SetKeyValue("THEMES", "POPULAR_MIXER", "1")
            '--------------------------------------------------------------------------------------
            oIni.SetKeyValue("KEYBOARD", "KB_ADD_01", "+")
            oIni.SetKeyValue("KEYBOARD", "KB_ADD_03", "N")
            oIni.SetKeyValue("KEYBOARD", "KB_DEL_01", "D")
            oIni.SetKeyValue("KEYBOARD", "KB_RET_01", ".")
            oIni.SetKeyValue("KEYBOARD", "KB_RST_01", "S")
            oIni.SetKeyValue("KEYBOARD", "KB_RST_03", "R")
            oIni.SetKeyValue("KEYBOARD", "KB_POP_01", "P")
            oIni.SetKeyValue("KEYBOARD", "KB_VIP_01", "V")
            oIni.SetKeyValue("KEYBOARD", "KB_PUP_01", "/")
            oIni.SetKeyValue("KEYBOARD", "KB_PDN_01", "*")
            oIni.SetKeyValue("KEYBOARD", "KB_VERIFY", "H")
            oIni.SetKeyValue("KEYBOARD", "KB_SWTPUB", "W")
            oIni.SetKeyValue("KEYBOARD", "KB__PAUSE", "L")
            oIni.SetKeyValue("KEYBOARD", "KB_SWTKAR", "K")

            '--------------------------------------------------------------------------------------
            oIni.Save(cini_path)

        End If
    End Sub

    Public Sub Save_Defa_Path2(Optional ByVal pRuta As String = "")
        If pRuta = "" Then
            pRuta = Application.StartupPath()
        End If
        If (bini_started = True) Then
            oIni.SetKeyValue("PATHS", "DIR_ODB", sgDir_odb)
            oIni.SetKeyValue("PATHS", "DIR_TMP", sgDir_Tmp)
            oIni.SetKeyValue("PATHS", "DIR_FL1", sgDir_Fls)
            oIni.SetKeyValue("PATHS", "DIR_FL2", sgDir_Fls2)
            oIni.SetKeyValue("PATHS", "DIR_IMG", sgDir_Img)
            oIni.SetKeyValue("PATHS", "DIR_MP3", sgDir_Mp3)
            oIni.SetKeyValue("PATHS", "DIR_PUB1", sgDir_Pub1)
            oIni.SetKeyValue("PATHS", "DIR_PUB2", sgDir_Pub2)
            oIni.SetKeyValue("PATHS", "FILE_BACKG", sgFle_Fon)
            '---------------------------------------------------------------------------------------------------------------
            oIni.SetKeyValue("TIMES", "TM_RET_GEN", Format(igDelay_Return_Gen, "#####0"))
            oIni.SetKeyValue("TIMES", "TM_RET_DIS", Format(igDelay_Return_Dis, "#####0"))
            oIni.SetKeyValue("TIMES", "TM_BON_VID", Format(igDelay_Bonus_Vid, "#####0"))
            oIni.SetKeyValue("TIMES", "ID_CNT_PRO", Format(sgIdx_Prm, "#####0"))
            '---------------------------------------------------------------------------------------------------------------
            oIni.SetKeyValue("ROCKOLA", "SHOW_VIDLA", Format(IIf(bgVideoLabel = True, 1, 0), "#####0"))
            oIni.SetKeyValue("ROCKOLA", "SHOW_DISLA", Format(IIf(bgDiscLabel = True, 1, 0), "#####0"))
            oIni.SetKeyValue("ROCKOLA", "SHOW_DISLA", Format(IIf(bgDiscLabel = True, 1, 0), "#####0"))
            oIni.SetKeyValue("ROCKOLA", "CR_ACC_KEY", Scramble(sgCr_AKey))
            oIni.SetKeyValue("ROCKOLA", "PASSW__BOX", Format(igShowPass, "#####0"))
            oIni.SetKeyValue("ROCKOLA", "START_MLST", Format(igStartPlayMusic, "#####0"))
            '---------------------------------------------------------------------------------------------------------------
            oIni.SetKeyValue("CREDITS", "LIMCR_CRED", Format(igLim_Cred, "#####0"))
            oIni.SetKeyValue("CREDITS", "KEEP_SCRED", Format(igKeep_Cred, "0"))
            oIni.SetKeyValue("CREDITS", "BONUS_CRED", Format(sgKb_BonC, "#####0"))
            oIni.SetKeyValue("CREDITS", "VIDEO_CRED", Format(sgKb_VID, "#####0"))
            oIni.SetKeyValue("CREDITS", "VIDEO_CRED", Format(sgKb_VID, "#####0"))
            oIni.SetKeyValue("CREDITS", "FLG_SAVECR", Format(igFlg_SavedCR, "#####0"))
            '---------------------------------------------------------------------------------------------------------------
            oIni.SetKeyValue("THEMES", "POPULAR_MIXER", Format(igMixe_Popu, "#####0"))
            '---------------------------------------------------------------------------------------------------------------
            oIni.SetKeyValue("KEYBOARD", "KB_ADD_01", sgKb_Crd1)
            oIni.SetKeyValue("KEYBOARD", "KB_ADD_03", sgKb_Crd2)
            oIni.SetKeyValue("KEYBOARD", "KB_DEL_01", sgKb_Del)
            oIni.SetKeyValue("KEYBOARD", "KB_RET_01", sgKb_Ret)
            oIni.SetKeyValue("KEYBOARD", "KB_RST_01", sgKb_ResM)
            oIni.SetKeyValue("KEYBOARD", "KB_RST_03", sgKb_ResA)
            oIni.SetKeyValue("KEYBOARD", "KB_POP_01", sgKb_Pop)
            oIni.SetKeyValue("KEYBOARD", "KB_VIP_01", sgKb_VIP)
            oIni.SetKeyValue("KEYBOARD", "KB_PUP_01", sgKb_UP)
            oIni.SetKeyValue("KEYBOARD", "KB_PDN_01", sgKb_DN)
            oIni.SetKeyValue("KEYBOARD", "KB_VERIFY", sgKb_Vef)
            oIni.SetKeyValue("KEYBOARD", "KB_SWTPUB", sgKb_SwP)
            oIni.SetKeyValue("KEYBOARD", "KB__PAUSE", sgKb_Pause)
            oIni.SetKeyValue("KEYBOARD", "KB_SWTKAR", sgKb_SwK)
            '---------------------------------------------------------------------------------------------------------------
            oIni.SetKeyValue("GENERAL", "SCRN_ALONE", Format(igScr_Alone, "#####0"))
            oIni.SetKeyValue("GENERAL", "SHOW_ONTOP", Format(bgKeep_On_Top, "#####0"))
            oIni.SetKeyValue("GENERAL", "NDUP_TEMES", Format(igNoDuplicT, "#####0"))
            oIni.SetKeyValue("GENERAL", "SWITCH_PUB", Format(IIf(bgSw_Pub = True, 1, 0), "#####0"))
            oIni.SetKeyValue("GENERAL", "SWITCH_KAR", Format(igInd_Kar, "#####0"))
            oIni.Save(cini_path)
        End If
    End Sub

    Public Sub Upd_Cnt()
        If (bini_started = True) Then
            oIni.SetKeyValue("GENERAL", "SWITCH_PUB", Format(IIf(bgSw_Pub = True, 1, 0), "#####0"))
            oIni.Save(cini_path)
        End If
    End Sub

    Public Sub Upd_Path(ByVal pRuta As String, ByRef pVariables() As String)
        If (bini_started = True) Then

            oIni.SetKeyValue("ROCKOLA", "FECHA_ACTI", Scramble(pVariables(7).ToString))
            oIni.SetKeyValue("ROCKOLA", "FECHA_ACTF", Scramble(pVariables(8).ToString))
            oIni.SetKeyValue("ROCKOLA", "SERIAL_CPU", Scramble(pVariables(14).ToString))
            oIni.SetKeyValue("ROCKOLA", "SERIAL_MAC", Scramble(pVariables(9).ToString))
            oIni.SetKeyValue("ROCKOLA", "NOMBRE_LOC", Scramble(pVariables(10).ToString))
            oIni.SetKeyValue("ROCKOLA", "WAPLIC_KEY", Scramble(pVariables(43).ToString))
            oIni.Save(cini_path)
        End If
    End Sub

    Public Sub Get_System_Path(ByRef paPath() As String, Optional ByVal pRuta As String = "")
        If pRuta = "" Then
            pRuta = Application.StartupPath()
        End If
        If (bini_started = True) Then
            paPath(1) = oIni.GetKeyValue("PATHS", "DIR_ODB", pRuta)
            paPath(2) = oIni.GetKeyValue("PATHS", "DIR_TMP", pRuta)
            paPath(3) = oIni.GetKeyValue("PATHS", "DIR_FL1", pRuta)
            sgDir_Fls2 = oIni.GetKeyValue("PATHS", "DIR_FL2", pRuta)
            paPath(4) = oIni.GetKeyValue("PATHS", "DIR_IMG", pRuta)
            paPath(5) = oIni.GetKeyValue("PATHS", "DIR_MP3", pRuta)
            paPath(6) = oIni.GetKeyValue("PATHS", "DIR_PUB1", pRuta)
            paPath(50) = oIni.GetKeyValue("PATHS", "DIR_PUB2", pRuta)
            paPath(13) = oIni.GetKeyValue("PATHS", "FILE_BACKG", "")
            '--------------------------------------------------------------------------------------
            paPath(7) = oIni.GetKeyValue("ROCKOLA", "FECHA_ACTI", "")
            paPath(7) = UnScramble(paPath(7))
            paPath(8) = oIni.GetKeyValue("ROCKOLA", "FECHA_ACTF", "")
            paPath(8) = UnScramble(paPath(8))
            paPath(9) = oIni.GetKeyValue("ROCKOLA", "SERIAL_MAC", "")
            paPath(9) = UnScramble(paPath(9))
            paPath(10) = oIni.GetKeyValue("ROCKOLA", "NOMBRE_LOC", "")
            paPath(10) = UnScramble(paPath(10))
            paPath(14) = oIni.GetKeyValue("ROCKOLA", "SERIAL_CPU", "")
            paPath(14) = UnScramble(paPath(14))
            paPath(43) = oIni.GetKeyValue("ROCKOLA", "WAPLIC_KEY", "")
            paPath(43) = UnScramble(paPath(43))
            paPath(44) = oIni.GetKeyValue("ROCKOLA", "SHOW_VIDLA", "0")
            paPath(45) = oIni.GetKeyValue("ROCKOLA", "SHOW_DISLA", "0")
            paPath(52) = oIni.GetKeyValue("ROCKOLA", "CR_ACC_KEY", "")
            paPath(52) = UnScramble(paPath(52))
            igShowPass = oIni.GetKeyValue("ROCKOLA", "PASSW__BOX", "0")
            igStartPlayMusic = Val(oIni.GetKeyValue("ROCKOLA", "START_MLST", "0"))
            '--------------------------------------------------------------------------------------
            paPath(46) = oIni.GetKeyValue("GENERAL", "SHOW_ONTOP", "0")
            paPath(47) = oIni.GetKeyValue("GENERAL", "SCRN_ALONE", "0")
            paPath(48) = oIni.GetKeyValue("GENERAL", "NDUP_TEMES", "0")
            paPath(49) = oIni.GetKeyValue("GENERAL", "SWITCH_PUB", "0")
            igInd_Kar = Val(oIni.GetKeyValue("GENERAL", "SWITCH_KAR", "0"))
            '--------------------------------------------------------------------------------------
            paPath(15) = oIni.GetKeyValue("TIMES", "TM_RET_GEN", "0")
            paPath(16) = oIni.GetKeyValue("TIMES", "TM_RET_DIS", "0")
            paPath(17) = oIni.GetKeyValue("TIMES", "TM_BON_VID", "20")
            paPath(53) = oIni.GetKeyValue("TIMES", "ID_CNT_PRO", "0")
            '--------------------------------------------------------------------------------------
            igFlg_SavedCR = Val(oIni.GetKeyValue("CREDITS", "FLG_SAVECR", "0"))

            paPath(20) = oIni.GetKeyValue("CREDITS", "LIMCR_CRED", "80")
            paPath(21) = oIni.GetKeyValue("CREDITS", "KEEP_SCRED", "0")
            paPath(38) = oIni.GetKeyValue("CREDITS", "BONUS_CRED", "0")
            paPath(36) = oIni.GetKeyValue("CREDITS", "VIDEO_CRED", "0")
            '--------------------------------------------------------------------------------------
            paPath(22) = oIni.GetKeyValue("THEMES", "POPULAR_MIXER", "0")
            '--------------------------------------------------------------------------------------
            paPath(23) = oIni.GetKeyValue("KEYBOARD", "KB_ADD_01", "+")
            paPath(25) = oIni.GetKeyValue("KEYBOARD", "KB_ADD_03", "N")
            paPath(27) = oIni.GetKeyValue("KEYBOARD", "KB_DEL_01", "D")
            paPath(29) = oIni.GetKeyValue("KEYBOARD", "KB_RET_01", ".")
            paPath(31) = oIni.GetKeyValue("KEYBOARD", "KB_RST_01", "S")
            paPath(33) = oIni.GetKeyValue("KEYBOARD", "KB_RST_03", "R")
            paPath(35) = oIni.GetKeyValue("KEYBOARD", "KB_POP_01", "P")
            paPath(37) = oIni.GetKeyValue("KEYBOARD", "KB_VIP_01", "V")
            paPath(39) = oIni.GetKeyValue("KEYBOARD", "KB_PUP_01", "/")
            paPath(41) = oIni.GetKeyValue("KEYBOARD", "KB_PDN_01", "*")
            paPath(40) = oIni.GetKeyValue("KEYBOARD", "KB_VERIFY", "H")
            paPath(51) = oIni.GetKeyValue("KEYBOARD", "KB_SWTPUB", "W")
            sgKb_Pause = oIni.GetKeyValue("KEYBOARD", "KB__PAUSE", "L")
            sgKb_SwK = oIni.GetKeyValue("KEYBOARD", "KB_SWTKAR", "K")
        End If
    End Sub

    Public Sub KeepMonitorActive()
        SetThreadExecutionState(EXECUTION_STATE.ES_DISPLAY_REQUIRED + EXECUTION_STATE.ES_CONTINUOUS) 'Do not Go To Sleep
    End Sub

    Public Function Scramble(Text As String) As String
        Dim i As Integer
        Dim c As Integer
        Dim Temp As String
        Temp = ""
        For i = 1 To Len(Text)
            c = Asc(Mid(Text, i, 1))
            c = c + 10

            'If the character was near the end of the ASCII character set then
            'adding 10 to the ASCII value can tip it over 255 which is the last
            'character - in this case, find the difference and that's the new
            'character (ie. at the start of the character set
            If c > 255 Then c = c - 255
            Temp = Temp & Chr(c)
        Next i
        Scramble = Temp
    End Function

    Public Function UnScramble(Text As String) As String
        Dim i As Integer
        Dim c As Integer
        Dim Temp As String
        Temp = ""
        For i = 1 To Len(Text)
            c = Asc(Mid(Text, i, 1))
            c = c - 10
            'If the character was near the start of the ASCII character set then
            'subtracting 10 from the ASCII value can take it under 0 which is the
            'first character - in this case, find the difference and that's the
            'new character (ie. at the start of the character set
            'You will notice that we add rather than subtract from 255 because if
            'you minus a minus it becomes a plus (!)
            If c < 0 Then c = 256 + c
            Temp = Temp & Chr(c)
        Next i
        UnScramble = Temp
    End Function

    Public Function GetTheComputerName() As String
        Dim ComputerName As String
        ComputerName = System.Net.Dns.GetHostName
        Return ComputerName
    End Function

    Public Function GetTheWindowsDirectory() As String
        GetTheWindowsDirectory = System.Environment.SystemDirectory
    End Function

    Public Function CheckValue(ByVal s As String) As String
        If VarType(s) = vbNull Then s = ""
        s = Trim(s)
        If s = "" Then CheckValue = "Not available" Else CheckValue = s
    End Function

    Public Sub ObPlayer_Ocupado(oPlayer As Object)
    End Sub

    Public Sub Sincronoze_Media()
    End Sub

    Public Sub ControlError()
    End Sub

    Public Sub Sleep(Seconds)
        Thread.Sleep(Seconds)
    End Sub

    Public Function PADL(sIn_Val As String, iIn_Len As Integer, sIn_Simbol As String)
        PADL = sIn_Val.PadLeft(iIn_Len, sIn_Simbol)
    End Function

    Public Function PADR(sIn_Val As String, iIn_Len As Integer, sIn_Simbol As String)
        PADR = sIn_Val.PadRight(iIn_Len, sIn_Simbol)
    End Function

    Function Proper(ByVal sIn_Val As String) As String
        Dim curCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
        Dim tInfo As TextInfo = curCulture.TextInfo()
        Proper = tInfo.ToTitleCase(sIn_Val)
    End Function

    Public Function GenerateRandomFileName() As String
        GenerateRandomFileName = Path.GetTempFileName()
    End Function

    Public Sub ControlPanels(filename As String)
        Dim rtn As Double
        On Error Resume Next
        rtn = Shell(filename, 5)
    End Sub

    Public Function GetWindowsDir() As String
        GetWindowsDir = System.Environment.SystemDirectory
    End Function

    Public Sub CenterForm(ByVal frm As Form, Optional ByVal parent As Form = Nothing)

        Dim x As Integer
        Dim y As Integer
        Dim r As Rectangle

        If Not parent Is Nothing Then
            r = parent.ClientRectangle
            x = r.Width - frm.Width + parent.Left
            y = r.Height - frm.Height + parent.Top
        Else
            r = Screen.PrimaryScreen.WorkingArea
            x = r.Width - frm.Width
            y = r.Height - frm.Height
        End If

        x = CInt(x / 2)
        y = CInt(y / 2)

        frm.StartPosition = FormStartPosition.Manual
        frm.Location = New Point(x, y)
    End Sub

    Private Function CpuId() As String
        Dim computer As String = "."
        Dim wmi As Object = GetObject("winmgmts:" &
        "{impersonationLevel=impersonate}!\\" &
        computer & "\root\cimv2")
        Dim processors As Object = wmi.ExecQuery("Select * from " &
        "Win32_Processor")

        Dim cpu_ids As String = ""
        For Each cpu As Object In processors
            cpu_ids = cpu_ids & ", " & cpu.ProcessorId
        Next cpu
        If cpu_ids.Length > 0 Then cpu_ids =
        cpu_ids.Substring(2)

        Return cpu_ids
    End Function


    Public Function MBCPUNumber() As String
        MBCPUNumber = CpuId()
    End Function

    Function VBRegSvr32(ByVal sServerPath As String, Optional fRegister As Boolean = True) As Boolean
        VBRegSvr32 = False
    End Function

    Public Sub SetCursorPosition(oObject As Object, xPos As Long, yPos As Long)
        oObject.Cursor = New Cursor(Cursor.Current.Handle)
        Cursor.Position = New Point(xPos, yPos)
        Cursor.Clip = New Rectangle(xPos, yPos, oObject.Width, oObject.Height)
    End Sub

    Public Sub Shuffle(ByRef data() As Integer, iStartIndex As Integer, iEndIndex As Integer)
        Dim x, swap As Integer
        Dim r As Random = New Random()

        For i = iStartIndex To iEndIndex
            x = r.Next(0, i)
            swap = data(x)
            data(x) = data(i)
            data(i) = swap
        Next i
    End Sub

    Public Sub Desordenar_array(ByRef aList() As Integer, startIndex As Integer, endIndex As Integer)
        Shuffle(aList, startIndex, endIndex)
    End Sub

    Function InList(cString As String, olist() As String)
        If (Array.IndexOf(olist, cString > -1)) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Refresh_Paginero(ByVal piPag_No As Integer, ByVal ipPag_Tot As Integer)
        If piPag_No = 1 And ipPag_Tot = 1 Then
            bgBlinkPag = False
            FMain.Img_PagPrev.Visible = False
            FMain.Img_PagNext.Visible = False
            'fMain.olPaginas.ForeColor = &HFFFF&
            'fMain.olPaginas2.ForeColor = &HFFFF&
        ElseIf piPag_No = 1 And ipPag_Tot > 1 Then
            bgBlinkPag = True
            FMain.Img_PagPrev.Visible = False
            FMain.Img_PagNext.Visible = True
        ElseIf piPag_No <> 1 And (ipPag_Tot = piPag_No) Then
            bgBlinkPag = True
            FMain.Img_PagPrev.Visible = True
            FMain.Img_PagNext.Visible = False
        ElseIf piPag_No > 1 And (ipPag_Tot <> piPag_No) Then
            bgBlinkPag = True
            FMain.Img_PagPrev.Visible = True
            FMain.Img_PagNext.Visible = True
        End If
    End Sub

    Public Function Shuffle_Return(ByVal items() As String) As Array
        Dim max_index As Integer = items.Length - 1
        Dim rnd As New Random(DateTime.Now.Millisecond)
        For i As Integer = 0 To max_index
            ' Pick an item for position i.
            Randomize()
            Dim j As Integer = rnd.Next(i, max_index)
            ' Swap them.
            Dim temp As String = items(i)
            items(i) = items(j)
            items(j) = temp
        Next i
        Return items
    End Function

    Sub Shuffle_ListBox(ByRef pListbox As System.Windows.Forms.ListBox, bLetFirst As Boolean)
        Dim cValue As String = ""
        If (pListbox.Items.Count > 0) Then
            cValue = pListbox.Items(0)
        End If

        Dim Random As New System.Random
        pListbox.BeginUpdate()
        If (bLetFirst = True) Then
            pListbox.Items.RemoveAt(0)
        End If
        Dim ArrayList As New System.Collections.ArrayList(pListbox.Items)
        pListbox.Items.Clear()
        While ArrayList.Count > 0
            Dim Index As System.Int32 = Random.Next(0, ArrayList.Count)
            pListbox.Items.Add(ArrayList(Index))
            ArrayList.RemoveAt(Index)
        End While
        If (bLetFirst = True) Then
            pListbox.Items.Insert(0, cValue)
        End If
        pListbox.EndUpdate()
    End Sub

    Sub Max_Pag_Gen()
        Dim pFileName = Application.StartupPath() & "\FilesV3.db;FailIfMissing=False;"
        Dim str As String = "DataSource=" & pFileName
        Dim con As SQLiteConnection
        Dim cmm As SQLiteCommand
        Dim dar As SQLiteDataReader
        Dim sql As String
        Try
            '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------
            con = New SQLiteConnection(str)

            If Not (con.State = ConnectionState.Open) Then
                con.Open()
            End If

            sql = "SELECT MAX(page) AS Max_PageG FROM file01 "
            cmm = New SQLiteCommand(sql, con)
            dar = cmm.ExecuteReader()

            '----------------------------------------------------------------------------------
            If dar.HasRows = True Then
                dar.Read()
                igMax_RgG = Val(dar.Item("Max_PageG").ToString())
            Else
                igMax_RgG = 0
            End If
            '----------------------------------------------------------------------------------

        Catch ex As SQLiteException
            Console.WriteLine("Error: " & ex.ToString())
        Finally
            dar.Close()
            dar = Nothing
            con.Close()
            con = Nothing
            cmm = Nothing
        End Try
    End Sub

    Sub Max_Pag_Dis(ByVal pId_Cod_Gen As String)
        '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------
        Dim pFileName = Application.StartupPath() & "\FilesV3.db;FailIfMissing=False;"
        Dim str As String = "DataSource=" & pFileName
        Dim con As SQLiteConnection
        Dim cmm As SQLiteCommand
        Dim dar As SQLiteDataReader
        Dim sql As String
        Try
            con = New SQLiteConnection(str)

            If Not (con.State = ConnectionState.Open) Then
                con.Open()
            End If

            sql = "SELECT MAX(page) as Max_PageD FROM file02 WHERE id_gen=" & pId_Cod_Gen & ""
            cmm = New SQLiteCommand(sql, con)
            dar = cmm.ExecuteReader()
            '----------------------------------------------------------------------------------
            If dar.HasRows = True Then
                dar.Read()
                igMax_RgD = Val(dar.Item("Max_PageD").ToString())
            Else
                igMax_RgD = 0
            End If
            '----------------------------------------------------------------------------------

        Catch ex As SQLiteException
            Console.WriteLine("Error: " & ex.ToString())
        Finally
            dar.Close()
            dar = Nothing
            con.Close()
            con = Nothing
            cmm = Nothing
        End Try
    End Sub

    Sub Max_Pag_Canc(ByVal pId_Cod_Disc As String)
        Dim pFileName = Application.StartupPath() & "\FilesV3.db;FailIfMissing=False;"
        Dim str As String = "DataSource=" & pFileName
        Dim con As SQLiteConnection
        Dim cmm As SQLiteCommand
        Dim dar As SQLiteDataReader
        Dim sql As String
        Try

            '------CONEXION A LA BASE DE DATOS SQL LITE----------------------------------------
            con = New SQLiteConnection(str)

            If Not (con.State = ConnectionState.Open) Then
                con.Open()
            End If

            sql = "SELECT MAX(page) AS Max_PageC FROM file03 WHERE id_dis=" & pId_Cod_Disc & ""
            cmm = New SQLiteCommand(sql, con)
            dar = cmm.ExecuteReader()
            '----------------------------------------------------------------------------------
            If dar.HasRows = True Then
                dar.Read()
                igMax_RgC = Val(dar.Item("Max_PageC").ToString())
            Else
                igMax_RgC = 0
            End If
            '----------------------------------------------------------------------------------
        Catch ex As SQLiteException
            Console.WriteLine("Error: " & ex.ToString())
        Finally

            dar.Close()
            dar = Nothing
            con.Close()
            con = Nothing
            cmm = Nothing
        End Try
    End Sub


End Module
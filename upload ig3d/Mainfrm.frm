VERSION 5.00
Begin VB.Form Mainfrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7080
   ClientLeft      =   1530
   ClientTop       =   1440
   ClientWidth     =   7650
   ControlBox      =   0   'False
   Icon            =   "Mainfrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTemp 
      Height          =   375
      Left            =   11880
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'################################################################################################################
'
'  Application Title    :       IG3D
'  Developer            :       Gaurang Vyas
'                               (079)-5469001
'                               gaurangvyas@hotmail.com
'                               gaurangjvyas@yahoo.com
'
'  File                 :       Mainfrm.frm
'  Content              :       Creates 3d Room
'
'################################################################################################################

Dim blnBlank As Boolean
Dim blnHalf1 As Boolean
Dim blnHalf2 As Boolean
Dim Oldx  As Single
Dim OldY As Single
Dim blnTry As Boolean
Dim intHDist As Integer
Dim intVDist As Integer
Public intHButtaDist As Integer
Public intVButtaDist  As Integer
Dim intButtaDist As Integer
Dim blnCrossButta As Boolean
Dim dbConnection As New ADODB.Connection
Dim rsTemplate As New ADODB.Recordset
Public DI As DirectInput
Public DIDev As DirectInputDevice
Dim FirstTileHeight  As Single
Dim blnFloorDone As Boolean
Dim intTileRemainingHeight  As Single
Dim MOUSESTATE As DIMOUSESTATE
Private strTempRandom As String
Private strTempStud As String
Dim FinalHeight As Single

Public Sub Start3dDemo()
Dim i As Single
Dim j As Single
Dim Length As Integer
Dim Width As Integer
Dim tilewidth As Single
Dim tileheight As Single
Dim wallheight As Single
Dim WallWidth As Single
Dim WallLength As Single
Dim intTileTypeId As Single
Dim OldHeight As Single
Dim intCrossRemain As Single
Dim Oldtilewidth As Single
Dim FirstTileWidth As Double

    Me.Show
    handle = Mainfrm.hwnd
    intRandomCount = 0
    UseHardWare = True
    Call InitWindow("", "IID_IDirect3DHALDevice")
    Set DINPUT = m_dx.DirectInputCreate
    Set DIdevice = DINPUT.CreateDevice("GUID_SysKeyboard")
    DIdevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIdevice.SetCooperativeLevel Me.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    Set DI = m_dx.DirectInputCreate
    Set DIDev = DI.CreateDevice("GUID_SYSMOUSE")
    Call DIDev.SetCommonDataFormat(DIFORMAT_MOUSE)
   
    intRed = 0
    intGreen = 0
    intBlue = 0
    
    Call m_rmFrameScene.SetSceneBackgroundRGB(intRed, intGreen, intBlue)

    Call CreateRoom3d(32, 84, 84, 8, 8, 2, 1, App.Path & "\tile1.bmp", 0)
    Call CreateRoom3d(32, 84, 84, 3, 8, 2, 1, App.Path & "\tile2.bmp", 29.5)
    Call CreateRoom3d(35 + 32, 84, 84, 8, 8, 2, 1, App.Path & "\tile3.bmp", 35)
    Call CreateRoom3d(35 + 32, 84, 84, 3, 8, 2, 1, App.Path & "\tile2.bmp", 35 + 29.5)
    Call LoadTexture(FloorTexture, App.Path & "\tile3.bmp", 256, 256)
    Call CreateFloorWall(84, 84, 8, 8, 8)
    blnMove = True
    Call SetWholeWorld
End Sub
Private Sub CreateRoom3d(RoomHeight As Single, RoomWidth As Single, RoomLength As Single, tileheight As Single, tilewidth As Single, TileType As Single, intTextureId As Single, ByVal strTileName As String, intYPos As Single, Optional intRemainingHeight As Single = 0)
Dim i As Single
Dim j As Single
Dim intNewTileHeight As Single
Dim intRemain As Single

blnMove = True
If Trim(strTileName) <> "" Then
    Call LoadTexture(TextureImage(intTextureId), strTileName, 128, 512)
End If
intNewTileHeight = tileheight / 2
   If intRemainingHeight <> 0 Then
      intRemain = 1 - (intRemainingHeight) / tileheight
   Else
      intRemainingHeight = tileheight
   End If
   For j = intYPos To RoomHeight - 1 Step tileheight
        For i = -tilewidth / 2 To RoomWidth + tilewidth Step tilewidth
        If blnCross = True Then
            If blnDiffCross = True Then
            Else
                Create3DVertexBackSide i, j, RoomLength, tilewidth, tileheight, 0, TextureImage(intTextureId)
                Create3DVertexBackSide i + intNewTileHeight, j - intNewTileHeight, RoomLength, tilewidth, tileheight, 0, TextureImage(intTextureId)
            End If
        Else
                Create3DVertexBackSide i, j, RoomLength, tilewidth, intRemainingHeight, 0, TextureImage(intTextureId), intRemain
        End If
        Next i
    Next j

    For j = intYPos To RoomHeight - 1 Step tileheight
      For i = RoomLength + intRemaining - tilewidth / 2 To -intRemaining - tilewidth Step -tilewidth
          If blnCross = True Then
                If blnDiffCross = True Then
                Else
                    Create3DVertexLeftSide RoomWidth, j, i, 0, tileheight, tilewidth, TextureImage(intTextureId)
                    Create3DVertexLeftSide RoomWidth, j - intNewTileHeight, i + intNewTileHeight, 0, tileheight, tilewidth, TextureImage(intTextureId)
                End If
          Else
                    Create3DVertexLeftSide RoomWidth, j, i, 0, intRemainingHeight, tilewidth, TextureImage(intTextureId), intRemain
          End If
       Next i
    Next j
    Dim intOld As Single
    intOld = RoomWidth - ((tilewidth) * Int(RoomWidth / tilewidth))
    For j = -(tilewidth / 2) To RoomWidth + tilewidth Step tilewidth
      For i = intYPos To RoomHeight - 1 Step tileheight
          If blnCross = True Then
            If blnDiffCross = True Then
            Else
                Create3DVertexFrontSide j, i, 0, tilewidth, tileheight, 0, TextureImage(intTextureId)
                Create3DVertexFrontSide j + intNewTileHeight, i - intNewTileHeight, 0, tilewidth, tileheight, 0, TextureImage(intTextureId)
            End If
          Else
                Create3DVertexFrontSide j, i, 0, tilewidth, intRemainingHeight, 0, TextureImage(intTextureId), intRemain
          End If
       Next i
    Next j

    For j = intYPos To RoomHeight - 1 Step tileheight
       For i = RoomLength + tilewidth / 2 To -tilewidth Step -tilewidth
          If blnCross = True Then
                If blnDiffCross = True Then
                Else
                    Create3DVertexRightSide 0, j, i, 0, tileheight, tilewidth, TextureImage(intTextureId)
                    Create3DVertexRightSide 0, j - intNewTileHeight, i - intNewTileHeight, 0, tileheight, tilewidth, TextureImage(intTextureId)
                End If
          Else
                    Create3DVertexRightSide 0, j, i, 0, intRemainingHeight, tilewidth, TextureImage(intTextureId), intRemain
          End If
       Next i
    Next j
End Sub

Public Sub SetWholeWorld()
    Px = (84 / 2)
    py = (84 / 3)
    RoomTotalWidth = 84
    RoomTotalHeight = 84
    RoomTotalLength = 84
    Pz = -(Math.Sqr((RoomTotalWidth * RoomTotalWidth) + (RoomTotalHeight * RoomTotalHeight)) - (RoomTotalLength / 2))
    blnOutSideView = True
    m_rmFrameCamera.SetPosition Nothing, Px, py, Pz
    blnMove = True
    Do While 1
        If blnMove = False Then
        Unload Me
        ShowCursor (1)
        Exit Sub
        End If
        RoomMovement
        DoEvents
        If blnMove = False Then
        Unload Me
        ShowCursor (1)
        Exit Sub
        End If
        UpdateRoomView
        DoEvents
    Loop
End Sub
Public Sub Create3DVertexBackSide(ByVal Px As Single, ByVal py As Single, ByVal Pz As Single, _
ByVal w As Single, ByVal h As Single, ByVal d As Single, Texture As Direct3DRMTexture3, Optional intPartial As Single = 0)
Dim fin As Direct3DRMFrame3
Dim tiLes As Single
Dim mout As Direct3DRMMeshBuilder3
Dim min As Direct3DRMMeshBuilder3
Dim f As Direct3DRMFace2
Dim frm As Direct3DRMFrame3

    tiLes = 1
    w = w / 2
    h = h / 2
    d = d / 2
    Set frm = m_rm.CreateFrame(Nothing)
    Set min = m_rm.CreateMeshBuilder
    min.SetQuality D3DRMRENDER_UNLITFLAT
    'Back
        Set f = m_rm.CreateFace
        With f
                        .AddVertex -w, h, d
                        .AddVertex w, h, d
                        .AddVertex w, -h, d
                        .AddVertex -w, -h, d
                        .SetTextureCoordinates 0, 1, intPartial
                        .SetTextureCoordinates 1, 0, intPartial
                        .SetTextureCoordinates 2, 0, 1
                        .SetTextureCoordinates 3, 1, 1
                        .SetTexture Texture
        End With
Out_Side1:
        min.AddFace f
        Set fin = m_rm.CreateFrame(frm)
        fin.AddVisual min
    frm.SetPosition Nothing, Px, py, Pz
    m_rmFrameScene.AddVisual frm
End Sub

Public Sub Create3DVertexLeftSide(ByVal Px As Single, ByVal py As Single, ByVal Pz As Single, _
ByVal w As Single, ByVal h As Single, ByVal d As Single, Texture As Direct3DRMTexture3, Optional intPartial As Single = 0)
    Dim tiLes As Single
    tiLes = 1
    w = w / 2
    h = h / 2
    d = d / 2
    Dim mout As Direct3DRMMeshBuilder3
    Dim min As Direct3DRMMeshBuilder3
    Dim f As Direct3DRMFace2
    Dim frm As Direct3DRMFrame3
    
    Set frm = m_rm.CreateFrame(Nothing)
    Set min = m_rm.CreateMeshBuilder
    min.SetQuality D3DRMRENDER_UNLITFLAT
    Set f = m_rm.CreateFace
    'Back
        With f
             .AddVertex w, -h, d
             .AddVertex w, h, d
             .AddVertex w, h, -d
             .AddVertex w, -h, -d
            .SetTextureCoordinates 0, 1, 1
            .SetTextureCoordinates 1, 1, intPartial
            .SetTextureCoordinates 2, 0, intPartial
            .SetTextureCoordinates 3, 0, 1
            .SetTexture Texture
       End With
Out_Side1:
        min.AddFace f
        Dim fin As Direct3DRMFrame3
        Set fin = m_rm.CreateFrame(frm)
        fin.AddVisual min
    frm.SetPosition Nothing, Px, py, Pz
 m_rmFrameScene.AddVisual frm
End Sub

Public Sub Create3DVertexRightSide(ByVal Px As Single, ByVal py As Single, ByVal Pz As Single, _
ByVal w As Single, ByVal h As Single, ByVal d As Single, Texture As Direct3DRMTexture3, Optional intPartial As Single = 0)
    Dim tiLes As Single
    tiLes = 1
    w = w / 2
    h = h / 2
    d = d / 2
    Dim mout As Direct3DRMMeshBuilder3
    Dim min As Direct3DRMMeshBuilder3
    Dim f As Direct3DRMFace2
    Dim frm As Direct3DRMFrame3
    
    Set frm = m_rm.CreateFrame(Nothing)
    Set min = m_rm.CreateMeshBuilder
    min.SetQuality D3DRMRENDER_UNLITFLAT
    Set f = m_rm.CreateFace
        With f
             .AddVertex -w, -h, -d
             .AddVertex -w, h, -d
             .AddVertex -w, h, d
             .AddVertex -w, -h, d
               .SetTextureCoordinates 0, 1, 1
                .SetTextureCoordinates 1, 1, intPartial
                .SetTextureCoordinates 2, 0, intPartial
                .SetTextureCoordinates 3, 0, 1
                .SetTexture Texture
        End With
Out_Side1:
        min.AddFace f
        Dim fin As Direct3DRMFrame3
        Set fin = m_rm.CreateFrame(frm)
        fin.AddVisual min
        Call fin.SetSceneBackgroundRGB(intRedBlank, intGreenBlank, intBlueBlank)
    frm.SetPosition Nothing, Px, py, Pz
 m_rmFrameScene.AddVisual frm
End Sub

Public Sub Create3DVertexFrontSide(ByVal Px As Single, ByVal py As Single, ByVal Pz As Single, _
ByVal w As Single, ByVal h As Single, ByVal d As Single, Texture As Direct3DRMTexture3, Optional intPartial As Single = 0)
    Dim tiLes As Single
    tiLes = 1
    w = w / 2
    h = h / 2
    d = d / 2
    Dim mout As Direct3DRMMeshBuilder3
    Dim min As Direct3DRMMeshBuilder3
    Dim f As Direct3DRMFace2
    Dim frm As Direct3DRMFrame3
    
    Set frm = m_rm.CreateFrame(Nothing)
    Set min = m_rm.CreateMeshBuilder
    min.SetQuality D3DRMRENDER_UNLITFLAT
    Set f = m_rm.CreateFace
        With f
            .AddVertex -w, -h, -d
            .AddVertex w, -h, -d
            .AddVertex w, h, -d
            .AddVertex -w, h, -d
            .SetTextureCoordinates 0, 0, 1
            .SetTextureCoordinates 1, 1, 1
            .SetTextureCoordinates 2, 1, intPartial
            .SetTextureCoordinates 3, 0, intPartial
            .SetTexture Texture
        End With
Out_Side1:
        min.AddFace f
        Dim fin As Direct3DRMFrame3
        Set fin = m_rm.CreateFrame(frm)
        fin.AddVisual min
        Call fin.SetSceneBackgroundRGB(intRedBlank, intGreenBlank, intBlueBlank)
    frm.SetPosition Nothing, Px, py, Pz
 m_rmFrameScene.AddVisual frm
End Sub

Public Sub Create3DVertexBottomSide(ByVal Px As Single, ByVal py As Single, ByVal Pz As Single, _
ByVal w As Single, ByVal h As Single, ByVal d As Single, Texture As Direct3DRMTexture3)
    Dim tiLes As Single
    tiLes = 1
    w = w / 2
    h = h / 2
    d = d / 2
    Dim mout As Direct3DRMMeshBuilder3
    Dim min As Direct3DRMMeshBuilder3
    Dim f As Direct3DRMFace2
    Dim frm As Direct3DRMFrame3
    
    Set frm = m_rm.CreateFrame(Nothing)
    Set min = m_rm.CreateMeshBuilder
    min.SetQuality D3DRMRENDER_UNLITFLAT
    Set f = m_rm.CreateFace
        With f
            .AddVertex -w, -h, d
            .AddVertex w, -h, d
            .AddVertex w, -h, -d
            .AddVertex -w, -h, -d
             .SetTextureCoordinates 0, 0, 0
            .SetTextureCoordinates 1, 1, 0
            .SetTextureCoordinates 2, 1, 1
            .SetTextureCoordinates 3, 0, 1
            .SetTexture Texture
        End With
        min.AddFace f
        Dim fin As Direct3DRMFrame3
        Set fin = m_rm.CreateFrame(frm)
        fin.AddVisual min
        Call fin.SetColorRGB(0, 0, 0)
    frm.SetPosition Nothing, Px, py, Pz
 m_rmFrameScene.AddVisual frm
End Sub

Private Sub cmdApplyColor_Click()
Dim intColor As Single
    cdlgColor.ShowColor
    intColor = cdlgColor.Color
    fraSelected.BackColor = intColor
    picAllObjects.BackColor = intColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set m_dd = Nothing
Set m_backBuffer = Nothing
Set m_ddClip = Nothing
Set m_dx = Nothing
Set m_frontBuffer = Nothing
Set FR_Building = Nothing
Set m_rmFrameCamera = Nothing
Set m_rmFrameScene = Nothing
Set m_rm = Nothing
Set m_rmDevice = Nothing
Set m_rm = Nothing
Set m_ddClip = Nothing
Set m_rmViewport = Nothing
Set DINPUT = Nothing
Set DIdevice = Nothing
Set TextureImage(1) = Nothing
Set TextureImage(2) = Nothing
End Sub

Public Sub LoadTexture(surf As Direct3DRMTexture3, file As String, w As Long, h As Long)
    Set surf = m_rm.CreateTextureFromSurface(LoadTextureFromBMP4(file, w, h))
    surf.SetName UCase(file)
    surf.GenerateMIPMap
    surf.SetCacheOptions 0, D3DRMTEXTURE_STATIC
End Sub

Public Sub CreateFloorWall(RoomLength As Single, RoomWidth As Single, tileheight As Single, tilewidth As Single, FirstTileHeight As Single)
blnCross = False
intNewTileHeight = tileheight / 2
For j = RoomLength + (tileheight / 2) To -tilewidth / 2 Step -tileheight
     For i = tilewidth / 2 To RoomWidth + tilewidth Step tilewidth
        Create3DVertexBottomSide i, 0, j, tilewidth, 0, tileheight, FloorTexture
       Next i
    Next j
End Sub

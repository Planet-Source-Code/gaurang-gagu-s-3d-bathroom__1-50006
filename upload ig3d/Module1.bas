Attribute VB_Name = "modAll3d"
'################################################################################################################
'
'  Application Title    :       IG3D
'  Developer            :       Gaurang Vyas
'                               (079)-5469001
'                               gaurangvyas@hotmail.com
'                               gaurangjvyas@yahoo.com
'
'  File                 :       Module1.bas
'  Contents             :       Contains Constants and procedures which Draws 3d Room
'                               and also handles room movement by keyboard
'
'################################################################################################################

Option Explicit
Const Rotate1 = 0.00001
Const Rotate2 = 0.0001
Const tiLes = 15

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Length As Single
Public Screen_Width As Single
Public tilewidth As Single
Public tileheight As Single
Public wallheight As Single
Public strTile As String
Public strBorder As String

Public blnDiffCross As Boolean
Public blnBorders  As Boolean
Public blnCross As Boolean
Public blnMove As Boolean
Public Random_Frame(0 To 200) As Direct3DRMFrame3
Public Extra_Frame(0 To 200) As Direct3DRMFrame3
Public FR_Building As Direct3DRMFrame3
Public MS_Building As Direct3DRMMeshBuilder3
Public blnLoaded As Boolean
Public handle As Long
'- direct x object
Public m_dx As New DirectX7

'- direct draw objects
Public m_dd As DirectDraw4
Public m_ddClip As DirectDrawClipper
Public m_frontBuffer As DirectDrawSurface4
Public m_backBuffer As DirectDrawSurface4

'- direct 3drm objects
Public m_rm As Direct3DRM3
Public m_rmDevice As Direct3DRMDevice3
Public m_rmViewport As Direct3DRMViewport2
Public m_rmFrameScene As Direct3DRMFrame3
Public m_rmFrameCamera As Direct3DRMFrame3

'- state
Public m_strDDGuid As String               'DirectDraw device guid
Public m_strD3DGuid As String              'Direct3DRM device guid
Public m_emptyrect As RECT
Public wndr As RECT

'Textures
Public TextureImage(1 To 500) As Direct3DRMTexture3
Public RandomTexture(1 To 4) As Direct3DRMTexture3
Public StrudsImage As Direct3DRMTexture3
Public FloorTexture As Direct3DRMTexture3
Public ExtraTexture(0 To 100) As Direct3DRMTexture3
Public CeilingTexture As Direct3DRMTexture3
Public WallTexture As Direct3DRMTexture3

'Direct 3d Lights
Public LT_Ambient As Direct3DRMLight
Public DINPUT As DirectInput
Public DIdevice As DirectInputDevice
Public Px As Double
Public py As Double
Public Pz As Double
Public RPx As Single
Public RPy As Single
Public Rpz As Single
Public blnRandom As Boolean
Public intObjectType As Integer

'FOR WALL
'FOR WHOLE ROOM
Public RoomTotalHeight As Single
Public RoomTotalWidth As Single
Public RoomTotalLength As Single
Public blnTopView  As Boolean
Public intRedBlank As Integer
Public intGreenBlank As Integer
Public intBlueBlank As Integer
Public intRed  As Single
Public intGreen  As Single
Public intBlue  As Single
Public RandomTileWidth As Single
Public RandomTileHeight As Single
Public intWall  As Single
Public intWallNo  As Single
Public intRandomWallNo As Single
Public ExtraTextureInfoPx(1 To 100) As Single
Public ExtraTextureInfoPy(1 To 100) As Single
Public ExtraTextureInfoPz(1 To 100) As Single
Public ExtraTextureInfoWallNo(1 To 100) As Single
Public ExtraTextureNameInfo(1 To 100) As String
Public intExtraCount As Single
Public intExtraObjMovement As Single
Public intRandomTileMov As Single
Public intRandomTileMovVer As Single
Public intRandomTileMovHor As Single

Public intCommonTileHeight(0 To 100) As Single
Public intCommonTileWidth(0 To 100) As Single

Public blnExtraTexture As Boolean
Public blnUpperHalf As Boolean
Public intExtraSelected As Single
Public blnDiagonalRandom As Boolean
'Public intRandomCount As Integer
Public intViewType As Single
Public intViewDist As Single
Public intYSpeed As Single
Public intXSpeed As Single
Public intZSpeed As Single
Public intRandomRemaining  As Single

Public intTotalRandomTiles As Integer
Public intTotalStudTiles As Integer
Public intTotalWallSections As Integer
Public FirstTileHeight As Single

Type RandomTile
    strRandomTileName As String
    strRandomTileType As String
    intRandomHeight As Integer
    intRandomWidth As Integer
    m_rmRandomFrame As Direct3DRMFrame3
    intRandomInfoPx As Single
    intRandomInfoPy As Single
    intRandomInfoPz As Single
    intWallNo As Single
End Type

Type StudsTile
    strStudTileName As String
    strStudTileType As String
    intStudHeight As Integer
    intStudWidth As Integer
    m_rmStudFrame As Direct3DRMFrame3
    intStudInfoPx As Single
    intStudInfoPy As Single
    intStudInfoPz As Single
    intWallNo As Single
End Type

Type WallSection
    TotalHeight As Single
    tileheight As Single
    tilewidth As Single
    TileType As String
    SectionType As String
End Type

Type ObjectsInformation
    ObjectHeight As Single
    
End Type

Public blnOutSideView  As Boolean
Public StudsTileInfo(1 To 1000) As StudsTile
Public RandomTilesInfo(1 To 1000) As RandomTile
Public BasicRandomTiles(1 To 100) As RandomTile
Public WallSectionInfo(0 To 10) As WallSection
Public TempRandomTilesInfo As RandomTile
Public CeilingColor As Long
Public ColorAboveTile As Long
Public intSection As Integer
Public intPrevSection As Integer
Public intSetQualityOfBathRoom As Integer
Dim intPreservedViewType As Single
Dim i As Integer

Public Angle As Single
Public Angle1 As Single
Const DEGPI = 3.14 * 2

Public Sub main()
    Load Mainfrm
    Mainfrm.Start3dDemo
End Sub

Public Function InitWindow(ddrawguid As String, d3dguid As String)
Dim ddsd As DDSURFACEDESC2
Dim intQualityType As Single
    intSetQualityOfBathRoom = 2
    Set m_rm = m_dx.Direct3DRMCreate()
    
    Set m_rmFrameScene = m_rm.CreateFrame(Nothing)
    Set m_rmFrameCamera = m_rm.CreateFrame(m_rmFrameScene)
        
    Set m_rmDevice = Nothing
    Set m_rmViewport = Nothing
    
    m_strDDGuid = ddrawguid
    m_strD3DGuid = d3dguid
    
    If d3dguid = "" Then m_strD3DGuid = "IID_IDirect3DRGBDevice"
    Set m_dd = m_dx.DirectDraw4Create(m_strDDGuid)
    
    m_dd.SetCooperativeLevel handle, DDSCL_NORMAL 'DDSCL_FULLSCREEN
    'm_dd.GetDisplayMode (ddsd)
    
    'm_dd.SetDisplayMode 800, 600, 16, 0, DDSDM_DEFAULT
    
    ddsd.lFlags = DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
'    ddsd.lMipMapCount = 1
    Set m_frontBuffer = m_dd.CreateSurface(ddsd)
    
    Set m_ddClip = m_dd.CreateClipper(0)
    m_ddClip.SetHWnd handle
    m_frontBuffer.SetClipper m_ddClip
    
    ResizeWindowedDevice (m_strD3DGuid)
    blnRandom = False
    m_rmDevice.SetQuality D3DRMRENDER_GOURAUD
    Select Case intSetQualityOfBathRoom
    Case 1
        m_rmDevice.SetTextureQuality D3DRMTEXTURE_LINEAR
    Case 2
        m_rmDevice.SetTextureQuality D3DRMTEXTURE_LINEARMIPLINEAR
    Case 3
        m_rmDevice.SetTextureQuality D3DRMTEXTURE_LINEARMIPNEAREST
    Case 4
        m_rmDevice.SetTextureQuality D3DRMTEXTURE_MIPLINEAR
    Case 5
        m_rmDevice.SetTextureQuality D3DRMTEXTURE_MIPNEAREST
    Case 6
        m_rmDevice.SetTextureQuality D3DRMTEXTURE_NEAREST
    End Select
    intViewDist = 1
End Function

Public Function ResizeWindowedDevice(d3dg As String)
    Dim memflags As Long
    Dim r As RECT
    Dim ddsd As DDSURFACEDESC2
    
    Call m_dx.GetWindowRect(handle, r)
    ddsd.lWidth = r.Right - r.Left
    ddsd.lHeight = r.Bottom - r.Top
    memflags = DDSCAPS_VIDEOMEMORY
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_3DDEVICE Or memflags
    Set m_backBuffer = m_dd.CreateSurface(ddsd)
    Set m_rmDevice = m_rm.CreateDeviceFromSurface(d3dg, m_dd, m_backBuffer, 0)
    Set m_rmViewport = m_rm.CreateViewport(m_rmDevice, m_rmFrameCamera, 0, 0, ddsd.lWidth, ddsd.lHeight)
    m_rmViewport.SetBack 1000
    ResizeWindowedDevice = True
End Function

Public Sub UpdateRoomView()
    m_rmViewport.Clear D3DRMCLEAR_ZBUFFER Or D3DRMCLEAR_TARGET
    m_rmViewport.Render m_rmFrameScene
    m_rmDevice.Update
    Call m_dx.GetWindowRect(handle, wndr)
    m_frontBuffer.Blt wndr, m_backBuffer, m_emptyrect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    m_frontBuffer.SetFontBackColor (1254255)
    m_frontBuffer.DrawText 50, 430, "MOVE FORWARD - UP ARROW KEY", False
    m_frontBuffer.DrawText 50, 450, "MOVE BACKWARD - DOWN ARROW KEY", False
    m_frontBuffer.DrawText 50, 470, "MOVE LEFT (ROTATE) - LEFT ARROW KEY", False
    m_frontBuffer.DrawText 50, 490, "MOVE RIGHT (ROTATE) - RIGHT ARROW KEY", False
    m_frontBuffer.DrawText 50, 510, "LOOK UP /DOWN - HOME / END KEYS", False
    m_frontBuffer.DrawText 50, 530, "TOP VIEW - F6 KEY ", False
    m_frontBuffer.DrawText 50, 550, "IF U LIKE THIS THEN DO NOT FORGET TO MAIL ME AT GAURANGVYAS@HOTMAIL.COM", False
    m_frontBuffer.DrawText 50, 570, "I WILL CONSIDER YOUR VOTE TO CREATE MY STRONG RESUME", False
    If blnTopView = True Then
        m_frontBuffer.DrawText 250, 150, "PRESS F7 TO GO BACK TO NORMAL VIEW", False
    End If
End Sub

Public Function LoadTextureFromBMP4(Sfile As String, Optional SWidth As Long, Optional Sheight As Long) As DirectDrawSurface4
    Dim cdib As New cBitmap
    Dim ddsd As DDSURFACEDESC2
    Dim dds As DirectDrawSurface4
    Dim hdcSurface As Long
    Dim TMPDIB As cBitmap
    
    cdib.CreateFromPicture LoadPicture(Sfile)
    Set TMPDIB = cdib.Resample(256, 256)
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.lMipMapCount = 3
    ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE Or DDSCAPS_MIPMAP 'Or DDSCAPS_OVERLAY
    'ddsd.ddsCaps.lCaps2 = DDSCAPS2_OP
    'DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    ddsd.ddpfPixelFormat.lFourCC = DDPF_FOURCC
    ddsd.lWidth = 256
    ddsd.lHeight = 256
    
    Set dds = m_dd.CreateSurface(ddsd)
    dds.restore
    
    hdcSurface = dds.GetDC
    TMPDIB.PaintPicture hdcSurface
    dds.ReleaseDC hdcSurface
    
    Set TMPDIB = Nothing
    
    Set LoadTextureFromBMP4 = dds
End Function

Public Function LoadTextureFromBMP4Object(Sfile As String, Optional SWidth As Long, Optional Sheight As Long) As DirectDrawSurface4
    Dim cdib As New cBitmap
    Dim ddsd As DDSURFACEDESC2
    Dim dds As DirectDrawSurface4
    Dim hdcSurface As Long
    Dim TMPDIB As cBitmap
    
    cdib.CreateFromPicture LoadPicture(Sfile)
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.ddsCaps.lCaps = DDSCAPS_TEXTURE
    ddsd.ddpfPixelFormat.lFourCC = DDPF_FOURCC
    ddsd.lWidth = 256
    ddsd.lHeight = 256
    
    Dim key As DDCOLORKEY
    key.low = 0
    key.high = 0
    
    Set dds = m_dd.CreateSurface(ddsd)
    dds.SetColorKey DDCKEY_SRCBLT, key
    
    hdcSurface = dds.GetDC
    Set TMPDIB = cdib.Resample(256, 256)
    TMPDIB.PaintPicture hdcSurface
    dds.ReleaseDC hdcSurface
    
    Set TMPDIB = Nothing
    
    Set LoadTextureFromBMP4Object = dds
End Function

Public Sub RoomMovement()

'On Error GoTo KEY_HANDLE_ERROR

Dim Height As Single
Dim Width As Single
Dim fin As Direct3DRMFrame3
Dim tiLes As Single
Dim mout As Direct3DRMMeshBuilder3
Dim min As Direct3DRMMeshBuilder3
Dim f As Direct3DRMFace2
Dim frm As Direct3DRMFrame3
Dim w As Single
Dim d As Single
Dim h As Single
Dim keyb As DIKEYBOARDSTATE

    DIdevice.Acquire
    DIdevice.GetDeviceStateKeyboard keyb
    intObjectType = 1
    
        If keyb.key(DIK_F7) Then
            If blnTopView = True Then
                blnTopView = False
            Else
                Px = (RoomTotalWidth / 2)
                py = (RoomTotalHeight / 3)
                Pz = -(Math.Sqr((RoomTotalWidth * RoomTotalWidth) + (RoomTotalHeight * RoomTotalHeight)) - (RoomTotalLength / 2))
                Angle1 = 0
                Angle = 0
                blnOutSideView = True
                m_rmFrameCamera.AddRotation D3DRMCOMBINE_REPLACE, 1, 0, 0, (0 * 3.14 / 180)
                m_rmFrameCamera.SetPosition Nothing, Px, py, Pz
                Exit Sub
            End If
            If blnTopView = False Then
                intViewType = intPreservedViewType
                Angle1 = 0
                Angle = 0
                On Error Resume Next
                m_rmFrameCamera.SetPosition Nothing, Px, py, Pz
                m_rmFrameCamera.AddRotation D3DRMCOMBINE_REPLACE, 1, 0, 0, (intViewType * 3.14 / 180)
                m_rmFrameCamera.SetPosition Nothing, Px, py, Pz
                Exit Sub
            End If
        End If
    If keyb.key(DIK_ESCAPE) Then
        Set TextureImage(1) = Nothing
        Set TextureImage(2) = Nothing
        Set m_dx = Nothing
        Set DINPUT = Nothing
        blnMove = False
    End If
       
    If keyb.key(DIK_F6) Then
            blnTopView = True
            intViewType = 90
            m_rmFrameCamera.SetPosition Nothing, RoomTotalWidth / 2, (RoomTotalHeight * 2), (RoomTotalLength / 2)
            m_rmFrameCamera.AddRotation D3DRMCOMBINE_REPLACE, 1, 0, 0, (intViewType * 3.14 / 180)
            m_rmFrameCamera.SetPosition Nothing, RoomTotalWidth / 2, (RoomTotalHeight * 2), (RoomTotalLength / 2)
    End If
    
    If keyb.key(DIK_LEFT) <> 0 Then
        If Px < 0 Then Exit Sub
        If Px > RoomTotalWidth Then Exit Sub
        If Pz < 0 Then Exit Sub
        If Pz > RoomTotalLength Then Exit Sub
         intWall = intWall - 0.2
         If intWall < 0 Then
            intWall = 6 + intWall
         End If
         
        Angle1 = Angle1 + 0.1
        If Angle > DEGPI Then
            Angle = 0
        End If
        
         Select Case intWall
            Case Is < 1.5
                            intWallNo = 1
            Case Is < 3
                            intWallNo = 2
            Case Is < 4.5
                            intWallNo = 3
            Case Is < 6
                            intWallNo = 4
         End Select
         m_rmFrameCamera.SetOrientation m_rmFrameCamera, -Rotate1, 0, Rotate2, 0, 1, 0
    End If
    
    If keyb.key(DIK_RIGHT) <> 0 Then
        If Px < 0 Then Exit Sub
        If Px > RoomTotalWidth Then Exit Sub
        If Pz < 0 Then Exit Sub
        If Pz > RoomTotalLength Then Exit Sub
        Angle1 = Angle1 - 0.1
        If Angle < 0 Then
            Angle = DEGPI - (-Angle)
        End If
         
         intWall = intWall + 0.1
         If intWall > 6 Then intWall = 0
         Select Case intWall
            Case Is < 1.5
                            intWallNo = 1
            Case Is < 3
                            intWallNo = 2
            Case Is < 4.5
                            intWallNo = 3
            Case Is < 6
                            intWallNo = 4
         End Select
         m_rmFrameCamera.SetOrientation m_rmFrameCamera, Rotate1, 0, Rotate2, 0, 1, 0
    End If
    If keyb.key(DIK_UP) <> 0 Then
         Angle = DEGPI - Angle1
         Px = Px + (Sin(Angle) * 1.2)
         Pz = Pz + (Cos(Angle) * 1.2)
         If blnOutSideView = True Then
            m_rmFrameCamera.SetPosition Nothing, Px, py, Pz
            If Pz > 0 Then blnOutSideView = False
         Exit Sub
         End If
         If Pz > RoomTotalLength - 1 Or Pz < 1 Then
            Px = Px - (Sin(Angle) * 1.2)
            Pz = Pz - (Cos(Angle) * 1.2)
         End If
         If Px > RoomTotalWidth Or Px < 1 Then
            Px = Px - (Sin(Angle) * 1.2)
            Pz = Pz - (Cos(Angle) * 1.2)
         End If
         m_rmFrameCamera.SetPosition Nothing, Px, py, Pz
    End If
    If keyb.key(DIK_DOWN) <> 0 Then
        Angle = DEGPI - Angle1
        Px = Px - (Sin(Angle) * 1.2)
        Pz = Pz - (Cos(Angle) * 1.2)
         If Pz > RoomTotalLength - 1 Or Pz < 1 Then
            Px = Px + (Sin(Angle) * 1.2)
            Pz = Pz + (Cos(Angle) * 1.2)
         End If
         If Px > RoomTotalWidth Or Px < 1 Then
            Px = Px + (Sin(Angle) * 1.2)
            Pz = Pz + (Cos(Angle) * 1.2)
         End If
        m_rmFrameCamera.SetPosition Nothing, Px, py, Pz
    End If
    If keyb.key(DIK_HOME) <> 0 Then
            intViewType = intViewType - intViewDist
            If intViewType > -90 Then
                m_rmFrameCamera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -(intViewDist * 3.14 / 180)
            Else
            intViewType = intViewType + intViewDist
            End If
            intPreservedViewType = intViewType
    End If
    If keyb.key(DIK_END) <> 0 Then
            intViewType = intViewType + intViewDist
            If intViewType < 90 Then
                m_rmFrameCamera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, (intViewDist * 3.14 / 180)
            Else
               intViewType = intViewType - intViewDist
            End If
            intPreservedViewType = intViewType
    End If
    If keyb.key(DIK_2) Then
        If py > RoomTotalHeight - 2 Then Exit Sub
         py = py + intYSpeed
         m_rmFrameCamera.SetPosition Nothing, Px, py, Pz
    End If
    If keyb.key(DIK_5) Then
        If py < 2 Then Exit Sub
         py = py - intYSpeed
         m_rmFrameCamera.SetPosition Nothing, Px, py, Pz
    End If
End Sub

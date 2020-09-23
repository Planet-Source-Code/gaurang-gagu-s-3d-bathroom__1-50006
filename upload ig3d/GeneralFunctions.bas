Attribute VB_Name = "MODGeneral"
  
  Public Const strAppName = "IG3D"
  Public TotalTile As Integer
  Public TileIndex As Integer
  Public iRoomHeight As Single
  Public iRoomWidth As Single
  Public iRoomLength As Single
  Public DB As New ADODB.Connection
  Public iUserType As Integer 'For User Checking
  Public rstUserCheck As New ADODB.Recordset
  Public blnUserStatus As Boolean
  Public DirPath As String 'For Tile Directory
  Public SelectedTileIndex(0 To 100) As Integer 'tile index
  Public SelectedTile(0 To 100) As String 'tile string
  Public UserName As String 'Logged user name
  Public UserPass As String 'Logged pass
  Public blnReset As Boolean
  Public BlankSectionColorName As Single
  Public ImagePath As String
  Public intRandomCount As Integer
  Public RandomTiles(1 To 100) As String
  Public iRandomTileCount As Integer
  Public iStudsCount As Integer
  Public iStuds(1 To 100) As String
  Public intTypeID As Integer
  
  Public strRandomTileToUse As String
  Public strStudToUse As String
  

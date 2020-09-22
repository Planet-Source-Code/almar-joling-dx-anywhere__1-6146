VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DX-Anywhere - Almar Joling"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   Icon            =   "mousemov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Run DirectX"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Tracking the mouse."
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program is developed by Almar Joling
'Please do not use it for commercial programs and/or source code 'books',
'databases, sites, or in any other way WITHOUT permission of the author.
'If you are going to use this code, could you please put my name somewhere?
'I'm just a 15 year (almost 16) that wants to help programmers, and show
'that teenagers can build cool things too. Thank you
'
'Use the code at your own risk.
'ajoling@theheadoffice.com
'Visit my multiplayer internet game at http://www3.ewebcity.com/stqw
'---------------------------------------------------------------------------

Option Explicit

'<Declarations that DirectX needs>
Dim DX As New DirectX7
Dim DD As DirectDraw7
Dim DDClipper As DirectDrawClipper
Dim RM As Direct3DRM3
Dim RMDevice As Direct3DRMDevice3
Dim RMViewport As Direct3DRMViewport2
Dim RootFrame As Direct3DRMFrame3
Dim LightFrame As Direct3DRMFrame3

Dim CameraFrame As Direct3DRMFrame3
Dim ObjectFrame As Direct3DRMFrame3
Dim UVFrame As Direct3DRMFrame3
Dim Light As Direct3DRMLight

Dim MeshBuilder As Direct3DRMMeshBuilder3
Dim Object As Direct3DRMMeshBuilder3
Dim xWidth As Long
Dim xHeight As Long
Dim Running As Boolean
Dim Finished As Boolean
Dim UseHwnd As Long
'</Declarations that DirectX needs>


'<Declarations to keep the query window on top of all other windows>
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_NOTOPMOST = -2
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
'</Declarations to keep the query window on top of all other windows>

'<Declarations to query the current mouse position>
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
'</Declarations to query the current mouse position>

'<Declarations to get the window Rectangle (size)>
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'</Declarations to get the window Rectangle (size)>

Private gStop As Boolean

Private Sub Command1_Click()
    Dim MousePT As POINTAPI
    Dim PrevWindow As Long
    Dim CurWindow As Long
    Dim X As Long
    Dim Y As Long
    Dim ClassName As String
    Dim RetValue As Long

    'Track mouse here
    If Command1.Caption = "Start" Then
     'Put window on top of all other windows
      SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
       Command1.Caption = "Stop"
        gStop = False
         PrevWindow = 0
        
        'Track until user stops
        Do
            'Stop tracking
            If gStop = True Then Exit Do
             'Get the current mouse position
             Call GetCursorPos(MousePT)
              X = MousePT.X
              Y = MousePT.Y
            
              'Get window under mouse
              CurWindow = WindowFromPoint(X, Y)
            
            If CurWindow <> PrevWindow Then
             ClassName = String$(256, " ")
              PrevWindow = CurWindow
               'Get the name and handle (hwnd) of the window
                RetValue = GetClassName(CurWindow, ClassName, 255)
                 ClassName = Left$(ClassName, InStr(ClassName, vbNullChar) - 1)
                'Display name
                  Label1.Caption = "The mouse is over " & ClassName & "--Hwnd-->" & CurWindow
                   Text1.Text = CurWindow
                End If
            
            DoEvents
            
        Loop
    
    Else
        'Change window back to normal
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
        Command1.Caption = "Start"
        'Stop the mouse tracking
        gStop = True
    End If
    
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Run DirectX" Then
Command2.Caption = "Stop DirectX"
UseHwnd = Text1.Text
'Initialize Retained Mode
InitRM
'Initialize Scene
InitScene App.Path + "\cube3.x"
'Start rendering
RenderLoop
Else
Running = False
Command2.Caption = "Run DirectX"
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CleanUp
    gStop = True
    End
End Sub

Private Sub InitRM()
Dim UseWindow As RECT
    
'Create Direct Draw From Current Display Mode
 Set DD = DX.DirectDrawCreate("")
  
  'Create new clipper object and associate it with a window'
   Set DDClipper = DD.CreateClipper(0)
    
    'Set the clipper to the chose handle
     DDClipper.SetHWnd UseHwnd
        
    'Get the rect of the window
     GetClientRect UseHwnd, UseWindow
      'Width = rectancle.right
       xWidth = UseWindow.Right
        'Height = rectangle.bottom
         xHeight = UseWindow.Bottom
          'Create the Retained Mode object
         Set RM = DX.Direct3DRMCreate()
        'Create the Retained Mode device to draw to
        Set RMDevice = RM.CreateDeviceFromClipper(DDClipper, "", xWidth, xHeight)
        RMDevice.SetQuality D3DRMRENDER_GOURAUD Or D3DRMLIGHT_ON Or D3DRMFILL_SOLID Or D3DRMSHADE_GOURAUD
        RMDevice.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY
        'RMDevice.SetShades 2
End Sub

Sub InitScene(sMesh As String)


    'Setup a scene graph with a camera light and object
    Set RootFrame = RM.CreateFrame(Nothing)
    Set CameraFrame = RM.CreateFrame(RootFrame)
    Set LightFrame = RM.CreateFrame(RootFrame)
    Set ObjectFrame = RM.CreateFrame(RootFrame)
    'position the camera and create the Viewport
    'provide the device thre viewport uses to render, the frame whose orientation and position
    'is used to determine the camera, and a rectangle describing the extents of the viewport
    CameraFrame.SetPosition Nothing, 0, 0, -10
    Set RMViewport = RM.CreateViewport(RMDevice, CameraFrame, 0, 0, xWidth, xHeight)
    
    
    'create a white light and hang it off the light frame
    Set Light = RM.CreateLight(D3DRMLIGHT_DIRECTIONAL, vbWhite) '&HFFFFFFFF)
    
    LightFrame.AddLight Light
    
    'For this sample we will load x files with geometry only
    'so create a meshbuilder object
    Set MeshBuilder = RM.CreateMeshBuilder()
    MeshBuilder.LoadFromFile sMesh, 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    
 

'------->To add a texture change the filename! (and uncomment)<--------

'    Dim Tex As Direct3DRMTexture3
'    Set Tex = RM.LoadTexture(App.Path + "\logo.bmp")
'    MeshBuilder.SetTexture Tex
    
    ObjectFrame.AddVisual MeshBuilder
    
    'Have the object rotating
    ObjectFrame.SetRotation Nothing, 1, 1, 1, 0.05
    ObjectFrame.AddScale D3DRMCOMBINE_BEFORE, 1, 1, 1
    
    
End Sub

Private Sub RenderLoop()
 On Local Error Resume Next
  Dim T1 As Long, T2 As Long
  Dim Delta As Single
  Running = True
  

    Do While Running = True
        'update
        Light.SetLinearAttenuation 0.2 * Rnd
        RM.Tick 1
        RMViewport.Clear D3DRMCLEAR_ALL    'clear the rendering surface rectangle described by the viewport
        RMViewport.Render RootFrame 'render to the device
     
        RMDevice.Update   'blt the image to the screen
        DoEvents
        
    Loop

     End Sub

Sub CleanUp()
    Running = False
    Exit Sub
    Set Light = Nothing
    Set MeshBuilder = Nothing
    Set Object = Nothing
    Set LightFrame = Nothing
    Set CameraFrame = Nothing
    Set ObjectFrame = Nothing
    Set RootFrame = Nothing
    Set RMDevice = Nothing
    Set DDClipper = Nothing
    Set RM = Nothing
    Set DD = Nothing
 
End Sub


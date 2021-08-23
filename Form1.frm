VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   FillColor       =   &H80000001&
   ForeColor       =   &H8000000F&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command17 
      Height          =   350
      Left            =   7530
      Picture         =   "Form1.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Desmarca Todos los Componentes"
      Top             =   3900
      Width           =   350
   End
   Begin VB.CommandButton cmdReferencias 
      Height          =   350
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Buscar Referencias en Proyectos"
      Top             =   3900
      Width           =   350
   End
   Begin VB.CommandButton cmdStart 
      Height          =   405
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Parar Compilación"
      Top             =   3840
      Width           =   405
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   71
      Top             =   7860
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14941
            MinWidth        =   14941
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   1411
            MinWidth        =   1411
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerStart 
      Interval        =   1000
      Left            =   8670
      Top             =   4410
   End
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   345
      Left            =   540
      TabIndex        =   37
      ToolTipText     =   "Hora de comienzo de la compilación"
      Top             =   3870
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      _Version        =   393216
      MouseIcon       =   "Form1.frx":0B4C
      Format          =   108134402
      CurrentDate     =   42584
   End
   Begin VB.PictureBox Orden 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3840
      Picture         =   "Form1.frx":0CA6
      ScaleHeight     =   225
      ScaleWidth      =   330
      TabIndex        =   35
      ToolTipText     =   "Orden de Compilación"
      Top             =   3510
      Width           =   330
   End
   Begin VB.PictureBox Picture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4770
      Picture         =   "Form1.frx":10E6
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   70
      Top             =   3540
      Width           =   240
   End
   Begin VB.CommandButton cmdCompatibilidad 
      Height          =   350
      Left            =   4230
      Picture         =   "Form1.frx":1AE8
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Establece compatibilidad de proyecto"
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdShutdown 
      Height          =   350
      Left            =   4230
      Picture         =   "Form1.frx":1DF2
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Puede apagar el equipo al finalizar la compilación"
      Top             =   3870
      Width           =   350
   End
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   8610
      Top             =   5430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAutocheck 
      Height          =   345
      Left            =   1950
      Picture         =   "Form1.frx":2134
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "AutoCheck!"
      Top             =   3870
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   8580
      Top             =   2790
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":28B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2FB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":310B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":345D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3E6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4189
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":44A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4D7D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer 
      Interval        =   2000
      Left            =   8670
      Top             =   6000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8670
      Top             =   4920
   End
   Begin VB.Frame Frame 
      Caption         =   "Conexiones VNC"
      Height          =   1335
      Left            =   7290
      TabIndex        =   63
      Top             =   6510
      Width           =   1905
      Begin VB.ListBox List 
         Height          =   1035
         Left            =   90
         TabIndex        =   64
         Top             =   240
         Width           =   1725
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   350
      Left            =   7515
      Picture         =   "Form1.frx":4ED7
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Carp. Server Local"
      Top             =   3480
      Width           =   350
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   2175
      Left            =   120
      TabIndex        =   62
      Top             =   4290
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   3836
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Max             =   32000
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command10 
      Height          =   350
      Left            =   7050
      Picture         =   "Form1.frx":51E1
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Carp. Cliente Local"
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdPararCompilacion 
      Height          =   350
      Left            =   3840
      Picture         =   "Form1.frx":54EB
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Parar Compilación"
      Top             =   3870
      Width           =   350
   End
   Begin VB.CommandButton cmdRegistrar 
      Height          =   350
      Left            =   8865
      Picture         =   "Form1.frx":5635
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Registrar al ALG01"
      Top             =   3480
      Width           =   350
   End
   Begin VB.CommandButton cmdSalvar 
      Height          =   350
      Left            =   8415
      Picture         =   "Form1.frx":5A39
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Salvar Lista de Proyectos"
      Top             =   3480
      Width           =   350
   End
   Begin VB.TextBox txtAvisar 
      Height          =   345
      Left            =   5100
      TabIndex        =   42
      ToolTipText     =   "Lista de PC's separadas por "";"""
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Height          =   350
      Left            =   7980
      Picture         =   "Form1.frx":5B83
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Abrir Proyecto/s"
      Top             =   3480
      Width           =   350
   End
   Begin VB.Frame Frame5 
      Caption         =   "Copia de Archivos"
      Height          =   1335
      Left            =   120
      TabIndex        =   59
      Top             =   6510
      Width           =   7125
      Begin VB.CheckBox chkCopiaCierra2 
         Height          =   195
         Left            =   5430
         TabIndex        =   52
         ToolTipText     =   "Copia y Cierra"
         Top             =   990
         Width           =   195
      End
      Begin VB.CheckBox chkCopiaCierra 
         Height          =   195
         Left            =   5430
         TabIndex        =   48
         ToolTipText     =   "Copia y Cierra"
         Top             =   630
         Width           =   195
      End
      Begin VB.CommandButton Command6 
         Caption         =   "SDP"
         Height          =   345
         Left            =   3720
         TabIndex        =   46
         ToolTipText     =   "SDP !"
         Top             =   540
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Height          =   345
         Left            =   6690
         Picture         =   "Form1.frx":5D07
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Abrir Destino"
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox txtCarpeta 
         Height          =   345
         Left            =   4230
         TabIndex        =   47
         Top             =   540
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.ComboBox cmbVersion 
         Height          =   315
         Left            =   750
         TabIndex        =   44
         Top             =   180
         Width           =   1245
      End
      Begin VB.TextBox txtActTesting 
         Height          =   345
         Left            =   60
         TabIndex        =   45
         Text            =   "\\alg01\d\Otros\ComponentesTesting\1.01.126"
         Top             =   540
         Width           =   3645
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   345
         Left            =   5670
         TabIndex        =   49
         Top             =   540
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Height          =   345
         Left            =   6690
         Picture         =   "Form1.frx":63B9
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Abrir Destino"
         Top             =   930
         Width           =   375
      End
      Begin VB.TextBox txtActClientes 
         Height          =   330
         Left            =   60
         TabIndex        =   51
         Text            =   "\\alg01\d\Otros\ActualizacionesClientes\1.01.126.SP1\bin"
         Top             =   930
         Width           =   5325
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Copiar"
         Height          =   345
         Left            =   5670
         TabIndex        =   53
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Versión:"
         Height          =   300
         Index           =   0
         Left            =   90
         TabIndex        =   60
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "BO's"
      Height          =   3405
      Left            =   2430
      TabIndex        =   56
      Top             =   30
      Width           =   2175
      Begin VB.CommandButton Command12 
         Height          =   255
         Left            =   1890
         Picture         =   "Form1.frx":6A6B
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Marca los componentes"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton Command11 
         Height          =   255
         Left            =   1890
         Picture         =   "Form1.frx":6BB5
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Desmarca los componentes"
         Top             =   390
         Width           =   255
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "AlgStdFunc.dll"
         Height          =   225
         Index           =   10
         Left            =   90
         TabIndex        =   19
         Top             =   2050
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "BOCereales.dll"
         Height          =   225
         Index           =   16
         Left            =   90
         TabIndex        =   12
         Top             =   300
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "BOGesCom.dll"
         Height          =   225
         Index           =   15
         Left            =   90
         TabIndex        =   13
         Top             =   550
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "BOProduccion.dll"
         Height          =   225
         Index           =   17
         Left            =   90
         TabIndex        =   14
         Top             =   800
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "BOFiscal.dll"
         Height          =   225
         Index           =   14
         Left            =   90
         TabIndex        =   15
         Top             =   1050
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "BOContabilidad.dll"
         Height          =   225
         Index           =   12
         Left            =   90
         TabIndex        =   16
         Top             =   1300
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "BOGeneral.dll"
         Height          =   225
         Index           =   13
         Left            =   90
         TabIndex        =   17
         Top             =   1550
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "BOSeguridad.dll"
         Height          =   225
         Index           =   11
         Left            =   90
         TabIndex        =   18
         Top             =   1800
         Width           =   1995
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Varios"
      Height          =   3405
      Left            =   7050
      TabIndex        =   58
      Top             =   30
      Width           =   2175
      Begin VB.CheckBox chkComponente 
         Caption         =   "AlgMobile.dll"
         Height          =   225
         Index           =   33
         Left            =   90
         TabIndex        =   33
         Top             =   1050
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "AlgInterop.dll"
         Height          =   225
         Index           =   32
         Left            =   90
         TabIndex        =   32
         Top             =   800
         Width           =   1995
      End
      Begin VB.CommandButton Command16 
         Height          =   255
         Left            =   1890
         Picture         =   "Form1.frx":6CFF
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Marca los componentes"
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton Command15 
         Height          =   255
         Left            =   1890
         Picture         =   "Form1.frx":6E49
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Desmarca los componentes"
         Top             =   390
         Width           =   255
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "ALGControls.ocx"
         Height          =   225
         Index           =   19
         Left            =   90
         TabIndex        =   31
         Top             =   550
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "PowerMaskControl.ocx"
         Height          =   225
         Index           =   18
         Left            =   90
         TabIndex        =   30
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DS's"
      Height          =   3405
      Left            =   4740
      TabIndex        =   57
      Top             =   30
      Width           =   2175
      Begin VB.CheckBox chkComponente 
         Caption         =   "DataAccess.dll"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   29
         Top             =   2550
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "DataShare.dll"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   28
         Tag             =   "i"
         Top             =   2300
         Width           =   1995
      End
      Begin VB.CommandButton Command14 
         Height          =   255
         Left            =   1890
         Picture         =   "Form1.frx":6F93
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Desmarca los componentes"
         Top             =   390
         Width           =   255
      End
      Begin VB.CommandButton Command13 
         Height          =   255
         Left            =   1890
         Picture         =   "Form1.frx":70DD
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Marca los componentes"
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "SPCereales.dll"
         Height          =   225
         Index           =   9
         Left            =   90
         TabIndex        =   27
         Top             =   2050
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "DSCereales.dll"
         Height          =   225
         Index           =   7
         Left            =   90
         TabIndex        =   20
         Top             =   300
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "DSGesCom.dll"
         Height          =   225
         Index           =   6
         Left            =   90
         TabIndex        =   21
         Top             =   550
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "DSProduccion.dll"
         Height          =   225
         Index           =   8
         Left            =   90
         TabIndex        =   22
         Top             =   800
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "DSFiscal.dll"
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   23
         Top             =   1050
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "DSContabilidad.dll"
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   24
         Top             =   1300
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "DSGeneral.dll"
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   25
         Top             =   1550
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "DSSeguridad.dll"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   26
         Top             =   1800
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proyectos"
      Height          =   3405
      Left            =   120
      TabIndex        =   55
      Top             =   30
      Width           =   2175
      Begin VB.CheckBox chkComponente 
         Caption         =   "ReportsGescom2.dll"
         Height          =   225
         Index           =   31
         Left            =   90
         TabIndex        =   11
         Top             =   3050
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "ReportsCereales2.dll"
         Height          =   225
         Index           =   30
         Left            =   90
         TabIndex        =   10
         Top             =   2800
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "ReportsGescom.dll"
         Height          =   225
         Index           =   29
         Left            =   90
         TabIndex        =   9
         Top             =   2550
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "ReportsCereales.dll"
         Height          =   225
         Index           =   25
         Left            =   90
         TabIndex        =   8
         Top             =   2300
         Width           =   1995
      End
      Begin VB.CommandButton Command9 
         Height          =   255
         Left            =   1890
         Picture         =   "Form1.frx":7227
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Desmarca los componentes"
         Top             =   390
         Width           =   255
      End
      Begin VB.CommandButton Command7 
         Height          =   255
         Left            =   1890
         Picture         =   "Form1.frx":7371
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Marca los componentes"
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "Inicio.exe"
         Height          =   225
         Index           =   27
         Left            =   90
         TabIndex        =   7
         Top             =   2050
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "Seguridad.dll"
         Height          =   225
         Index           =   21
         Left            =   90
         TabIndex        =   6
         Top             =   1800
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "AdministradorGeneral.dll"
         Height          =   225
         Index           =   20
         Left            =   90
         TabIndex        =   5
         Top             =   1550
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "Contabilidad.dll"
         Height          =   225
         Index           =   26
         Left            =   90
         TabIndex        =   4
         Top             =   1300
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "Fiscal.dll"
         Height          =   225
         Index           =   24
         Left            =   90
         TabIndex        =   3
         Top             =   1050
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "Produccion.dll"
         Height          =   225
         Index           =   28
         Left            =   90
         TabIndex        =   2
         Top             =   800
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "GestionComercial.dll"
         Height          =   225
         Index           =   22
         Left            =   90
         TabIndex        =   1
         Top             =   550
         Width           =   1995
      End
      Begin VB.CheckBox chkComponente 
         Caption         =   "Cereales.dll"
         Height          =   225
         Index           =   23
         Left            =   90
         TabIndex        =   0
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.ComboBox cmbCarpeta 
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Top             =   3510
      Width           =   3705
   End
   Begin VB.ListBox lst1 
      BackColor       =   &H00808000&
      Height          =   2205
      ItemData        =   "Form1.frx":74BB
      Left            =   330
      List            =   "Form1.frx":74BD
      TabIndex        =   43
      Top             =   4290
      Width           =   8895
   End
   Begin VB.CommandButton cmdCompilar 
      BackColor       =   &H00808000&
      Caption         =   "Compilar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Compilar proyectos seleccionados"
      Top             =   3870
      Width           =   1515
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   375
      Left            =   8850
      OleObjectBlob   =   "Form1.frx":74BF
      SourceDoc       =   "D:\OGG\LOAT.wav"
      TabIndex        =   61
      Top             =   3870
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declaración del Api SetLayeredWindowAttributes que establece _
 la transparencia al form
  
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long
  
  
'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                 ByVal nIndex As Long) As Long
  
  
'Declaración del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes
  
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
  
  
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Color progress bar
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_USER = &H400
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR
'Color progress bar

'Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Private Const SND_ASYNC = &H1         '  play asynchronously

Private strAction         As String

Private strCarpetaTrabajo As String
Private strErrores        As String
Private strOutdir         As String

Private Const EJECUTABLE_VB6_X86       As String = "C:\Archivos de programa\Microsoft Visual Studio\VB98\vb6.exe"
Private Const COMPILADOR_VB6_X86       As String = "C:\Archivos de programa\Microsoft Visual Studio\VB98\vb6c.exe"

Private Const EJECUTABLE_VB6_X64       As String = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6.exe"
Private Const COMPILADOR_VB6_X64       As String = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\vb6c.exe"

Private Const ACTUALIZACIONES_CLIENTES_ACTUAL   As String = "\\Alg01\d\otros\SoftCerealBin\VersionActual\bin"
Private Const ACTUALIZACIONES_CLIENTES_ANTERIOR As String = "\\Alg01\d\otros\SoftCerealBin\VersionAnterior\bin"

Private Const OUTDIR      As String = "C:\Archivos de programa\Algoritmo"
Private Const OUTDIR_X86  As String = "C:\Program Files (x86)\Algoritmo"

Private Const COMPONENTES_TESTING      As String = "\\alg01\d\Otros\ComponentesTesting\1.01.126"

Private fs                As Object
   
Private ix                As Integer
Private iz                As Integer
Private bTesting          As Boolean
Private elapsed           As Long
Private aParametros()     As String
Private aEstadisticas(33) As Integer

Private PID               As Long
Private lngID_Vss         As Long
Private lngID_RAR         As Long
Private bHayErrores       As Boolean
Private bTerminoRAR       As Boolean

Private bBusqueSS         As Boolean
Private strSS             As String
Private iTiempo           As Integer
Private iTiempoTotal      As Integer
Private iTotDLLS          As Integer
Private bCompCancelada    As Boolean
Private bCompCerrado      As Boolean
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
'Private Declare Sub CloseHandle Lib "Kernel32" (ByVal hPass As Long)

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'-----------------
'CargarUsuario
Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type
Private Const ERROR_SUCCESS            As Long = 0
Private Const MIB_TCP_STATE_CLOSED     As Long = 1
Private Const MIB_TCP_STATE_LISTEN     As Long = 2
Private Const MIB_TCP_STATE_SYN_SENT   As Long = 3
Private Const MIB_TCP_STATE_SYN_RCVD   As Long = 4
Private Const MIB_TCP_STATE_ESTAB      As Long = 5
Private Const MIB_TCP_STATE_FIN_WAIT1  As Long = 6
Private Const MIB_TCP_STATE_FIN_WAIT2  As Long = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT As Long = 8
Private Const MIB_TCP_STATE_CLOSING    As Long = 9
Private Const MIB_TCP_STATE_LAST_ACK   As Long = 10
Private Const MIB_TCP_STATE_TIME_WAIT  As Long = 11
Private Const MIB_TCP_STATE_DELETE_TCB As Long = 12
Private Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal bcount As Long)
Private Declare Function lstrcpyA Lib "kernel32" (ByVal retVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Private Declare Function ntohs Lib "wsock32.dll" (ByVal addr As Long) As Long
'-----------------
'Recuperar_Nombre_Host
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    imaxsockets As Integer
    imaxudp As Integer
    lpszvenderinfo As Long
End Type
  
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal VersionReq As Long, WSADataReturn As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal S As String) As Long
Private Declare Function gethostbyaddr Lib "wsock32.dll" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
'-----------------
Private strNombreUsuario As String
Private bnewells As Boolean
Private Const CANTIDAD_TOTAL_DLL As Long = 33

Private Const BM_SETSTATE = &HF3 'Shutdown
Private Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long         'Shutdown
Private bShutdown As Boolean 'Shutdown

'-ServiceCommand
'Constantes para usar con OpenSCManager
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Enum SERVICE_CONTROL
    SERVICE_CONTROL_STOP = 1&
    SERVICE_CONTROL_PAUSE = 2&
    SERVICE_CONTROL_CONTINUE = 3&
    SERVICE_CONTROL_INTERROGATE = 4&
    SERVICE_CONTROL_SHUTDOWN = 5&
End Enum
'-
Private strReferencias      As String
Private Xmove               As Double
Private Const LB_SETHORIZONTALEXTENT = &H194
Private lLength As Long
Private lLengtext As Long

Private Sub chkComponente_Click(Index As Integer)
   
   TiempoNecesario Index, chkComponente(Index).Value
   If chkComponente(Index).Value = vbUnchecked Then Exit Sub
   
   StatusBar.Panels(1).Text = "Comp. Binaria: " & BuscarCompatibilidad(Index)
End Sub

Private Sub cmdPararCompilacion_Click()
   If KillProcess(PID, 0) Then
       Timer1.Enabled = False
       lst1.AddItem "Compilación Cancelada"
       lst1.ListIndex = lst1.ListCount - 1
       PID = 0
       bCompCancelada = True
   End If
End Sub

Private Sub cmdReferencias_Click()
   lst1.Clear
   For ix = 0 To CANTIDAD_TOTAL_DLL
      If chkComponente(ix) = vbChecked Then
         strReferencias = ""
         
         For iz = 0 To CANTIDAD_TOTAL_DLL
            BuscarReferencias chkComponente(ix).Index, chkComponente(iz).Index
         Next
         If Len(strReferencias) = 0 Then
            lst1.Clear

            lst1.ListIndex = lst1.ListCount - 1
         End If
      End If
   Next
End Sub

Private Sub cmdShutdown_Click()
Static Shutdown As Long

Shutdown = Not Shutdown

If Shutdown Then
    cmdShutdown.Picture = ImageList.ListImages(4).Picture
    cmdShutdown.ToolTipText = "El equipo se apagará al finalizar la compilación"
    bShutdown = True
Else
    cmdShutdown.Picture = ImageList.ListImages(3).Picture
    cmdShutdown.ToolTipText = "Puede apagar el equipo al finalizar la compilación"
    bShutdown = False
End If

'Call SendMessageBynum(cmdShutdown.hwnd, BM_SETSTATE, Shutdown, 0)
End Sub

Private Sub Command1_Click()
   Shell "explorer.exe " & GetOutdir & "\server", vbMaximizedFocus
End Sub

Private Sub Form_Load()
   Dim cnn        As ADODB.Connection
   Dim rst        As ADODB.Recordset
   Dim SQL        As String
   Dim parametros As String
   
   PID = El_pid("compilador.exe")
   If PID > 0 Then
      SetPriority PID, Below_Normal
   End If
   
   Call Aplicar_Transparencia(Me.hwnd, CByte(240))
   
   Revisar_checks
   
   Form1.Caption = "Compilador - Algoritmo S.A."
   
   Set cnn = New ADODB.Connection
   cnn.ConnectionString = "Provider=OraOLEDB.Oracle.1;Password=apfrms2001;User ID=SYSADMIN;Data Source=BASE"
   cnn.Open
   
   Set rst = New ADODB.Recordset
   rst.CursorLocation = adUseClient
   rst.LockType = adLockReadOnly
   rst.CursorType = adOpenStatic

   SQL = " SELECT DISTINCT VERSION_PRODUCTO "
   SQL = SQL & "  FROM SCRIPTS "
   SQL = SQL & "  ORDER BY VERSION_PRODUCTO DESC "

   rst.Open SQL, cnn
   
   For ix = 1 To 2
      cmbVersion.AddItem Left(rst("VERSION_PRODUCTO"), 1) & "." & Mid(rst("VERSION_PRODUCTO"), 2, 2) & "." & Mid(rst("VERSION_PRODUCTO"), 4, 3)
      rst.MoveNext
   Next
   cmbVersion.ListIndex = 0
   rst.Close
   
   Cargar_cmbCarpeta
   
   txtActTesting.Text = COMPONENTES_TESTING
   Select Case NombrePC
      Case "PC3"
         txtActClientes.Text = ACTUALIZACIONES_CLIENTES_ACTUAL
         txtActClientes.Enabled = True
         
         txtActTesting.Text = vbNullString
         txtActTesting.Enabled = False
         
         Command3.Enabled = False
         Command4.Enabled = False
         Command6.Enabled = False
         chkCopiaCierra.Enabled = False
         cmdShutdown.Enabled = False
         cmdRegistrar.Enabled = False
         cmbVersion.Enabled = False
         
         Command2.Enabled = True
         Command8.Enabled = True
         
         chkCopiaCierra2.Enabled = True
      Case "PC5"
         txtActClientes.Text = ACTUALIZACIONES_CLIENTES_ANTERIOR
         txtActClientes.Enabled = True
         
         txtActTesting.Text = vbNullString
         txtActTesting.Enabled = False
         
         Command3.Enabled = False
         Command4.Enabled = False
         Command6.Enabled = False
         chkCopiaCierra.Enabled = False
         cmdShutdown.Enabled = False
         cmdRegistrar.Enabled = False
         cmbVersion.Enabled = False
         
         Command2.Enabled = True
         Command8.Enabled = True
         
         chkCopiaCierra2.Enabled = True
      Case Else
         txtActClientes.Text = vbNullString
         txtActClientes.Enabled = False
         
         Command3.Enabled = True
         Command4.Enabled = True
         Command6.Enabled = True
         chkCopiaCierra.Enabled = True
         cmdShutdown.Enabled = True
         cmdRegistrar.Enabled = True
         
         Command2.Enabled = False
         Command8.Enabled = False
         
         chkCopiaCierra2.Enabled = False
   End Select
   

   
   cnn.Close
   
   'Parametros
   parametros = command$

   aParametros = Split(parametros, " ")
   For ix = LBound(aParametros) To UBound(aParametros)
      If ix = 0 Then
         Select Case aParametros(ix)
            Case "/?"
               MsgBox "<Compilador> compila los proyectos guardados en el archivo compilador.dat" & _
               " " & _
               "Uso: Compilador [compilador.dat]"
               End
            Case "compilador.dat"
               Revisar_checks
               For iz = 0 To CANTIDAD_TOTAL_DLL
                  If chkComponente(iz) = vbChecked Then
                     Compilar (iz)
                  End If
               Next
               'Enviar_Mail
               End
         End Select
      End If
   Next ix
   
   cmbVersion_Click
   SetBackColor ProgressBar.hwnd, &H808000
   LeerEstadisticas
   SetSourceSafe
   'CargarUsuario False
   
   If NombrePC = "ADRIAN" Or NombrePC = "JUANJO" Or NombrePC = "TEST1" Or (NombrePC = "PC3" And UCase(txtAvisar.Text) = "ADRIAN") Then
      cmdCompatibilidad.Enabled = True
   Else
      cmdCompatibilidad.Enabled = False
   End If
   
   cmdStart.Picture = ImageList.ListImages(7).Picture
   DTPicker.Enabled = False
   
   cmdReferencias.Picture = ImageList.ListImages(8).Picture
   
   'Inicia el servicio mensajero
   On Error Resume Next
   ServiceCommand "Messenger", 0
 
End Sub

Private Sub cmdCompilar_Click()
         Dim strLista As String
   
10       On Error GoTo ErrorHandler
      
20       lst1.Clear
30       DeshabilitarControles
31       LeerEstadisticas
         bHayErrores = False
         bBusqueSS = False
   
35       VerDllHosts
         
40       For ix = 0 To CANTIDAD_TOTAL_DLL
50          If chkComponente(ix) = vbChecked Then
60             Compilar (ix)
               If cmdCompatibilidad.Picture = ImageList.ListImages(6).Picture Then 'Compatibilidad de proyecto
                  For iz = 1 To lst1.ListCount
                     If InStr(lst1.List(iz), "Error") > 0 Or InStr(lst1.List(iz), "abierto") > 0 Then
                        bHayErrores = True
                     End If
                  Next
                  If bHayErrores Then
                     Exit For
                  End If
               End If
70          End If
80       Next
   
91       For ix = 1 To lst1.ListCount
92          If InStr(lst1.List(ix), "Error") > 0 Or InStr(lst1.List(ix), "abierto") > 0 Then
93             lst1.BackColor = &HFF&
94             bHayErrores = True
95          End If
96       Next
         
         'Call sndPlaySound("D:\OGG\LOAT.wav", SND_ASYNC)
100      'OLE1.Action = 7
   
110      strLista = ""
120      If Len(txtAvisar.Text) > 0 Then
130         For ix = 0 To CANTIDAD_TOTAL_DLL
140            If chkComponente(ix) = vbChecked Then
150               strLista = strLista & chkComponente(ix).Caption & " "
160            End If
170         Next
180         Avisar (strLista)
190      End If
         'Enviar_Mail
         
         If Not bHayErrores Then
191         SalvarEstadisticas
            ActualizarLog
         End If
         
192      If chkCopiaCierra.Value = vbChecked And Not bHayErrores Then
            Command3_Click
            If bShutdown Then
               Shell "shutdown -s -f -t 00"
            End If
            End
         End If
         
193      If chkCopiaCierra2.Value = vbChecked And Not bHayErrores Then
            Command2_Click 'Solo PC3/PC5
            End
         End If
         
194      If bShutdown Then
            Shell "shutdown -s -f -t 00"
         End If
         
90       HabilitarControles

200      Exit Sub
   
ErrorHandler:
210      strAction = "cmdCompilar_Click " & Err.Description & Erl
220      lst1.AddItem strAction

End Sub
Private Sub Compilar(ix As Integer)
         Dim ts              As TextStream
         Dim strEjecuta      As String
         Dim strDll          As String
         Dim strError        As String
         
10       On Error GoTo ErrorHandler
         
20       Set fs = CreateObject("Scripting.FileSystemObject")
         
30       elapsed = 0
40       lst1.BackColor = &H808000
         bCompCancelada = False
         bCompCerrado = False
         
41       ProgressBar.Max = IIf(aEstadisticas(ix) = 0, 1, aEstadisticas(ix))
42       If aEstadisticas(ix) = 1 Then
43          SetBarColor ProgressBar.hwnd, &H808000
44       Else
45          SetBarColor ProgressBar.hwnd, &HFF00&
46       End If
         
50       strAction = "Compilando " & chkComponente(ix).Caption & "..."
60       lst1.AddItem strAction
61       lst1.ListIndex = lst1.ListCount - 1
         lLengtext = 0
         getlistboxHScrollBar strAction
         
70       strDll = Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & "\" & Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & ".vbp"
'80       If InStr(chkComponente(ix).Caption, "PowerMask") Then
'90          strDll = Replace(strDll, "PowerMaskControl\", "PowerMask\")
'100      End If
         
         strCarpetaTrabajo = GetCarpetaTrabajo(chkComponente(ix).Caption)
         
110      If InStr(chkComponente(ix).Caption, "DS") Or _
            InStr(chkComponente(ix).Caption, "SP") Or _
            InStr(chkComponente(ix).Caption, "DataAccess") Or _
            InStr(chkComponente(ix).Caption, "DataShare") Then
160            strErrores = GetOutdir & "\server\Errores.txt"
170            strOutdir = GetOutdir & "\server"
         Else
            If InStr(chkComponente(ix).Caption, "AlgInterop") Or InStr(chkComponente(ix).Caption, "AlgMobile") Then
               strErrores = GetOutdir & "\services\Errores.txt"
               strOutdir = GetOutdir & "\services"
               If Not fs.FolderExists(strOutdir) Then
                  fs.CreateFolder (strOutdir)
               End If
            Else
130            strErrores = GetOutdir & "\Errores.txt"
140            strOutdir = GetOutdir
            End If
         End If
         
410      If fs.FileExists(strErrores) Then fs.DeleteFile (strErrores)
         
         'Cambiar proyecto
         If cmdCompatibilidad.Picture = ImageList.ListImages(6).Picture Then
            SetCompProyON strCarpetaTrabajo, strDll
         End If
         
420      If fs.FileExists(EJECUTABLE_VB6_X86) Then
430         If Not fs.FileExists(COMPILADOR_VB6_X86) Then
440            FileCopy EJECUTABLE_VB6_X86, COMPILADOR_VB6_X86
450         End If
460         strEjecuta = "cmd.exe /c start " & Chr(34) & "Compilando..." & Chr(34) & " /belownormal " & Chr(34) & _
                       COMPILADOR_VB6_X86 & Chr(34) & " /m " & Chr(34) & strCarpetaTrabajo & strDll & Chr(34) & _
                       " /out " & Chr(34) & strErrores & Chr(34) & " /outdir " & Chr(34) & strOutdir & Chr(34)

470      Else
480         If Not fs.FileExists(COMPILADOR_VB6_X64) Then
490            FileCopy EJECUTABLE_VB6_X64, COMPILADOR_VB6_X64
500         End If
510         strEjecuta = "cmd.exe /c start " & Chr(34) & "Compilando..." & Chr(34) & " /belownormal " & Chr(34) & _
                       COMPILADOR_VB6_X64 & Chr(34) & " /m " & Chr(34) & strCarpetaTrabajo & strDll & Chr(34) & _
                       " /out " & Chr(34) & strErrores & Chr(34) & " /outdir " & Chr(34) & strOutdir & Chr(34)

520      End If

530      If Not fs.FileExists(strCarpetaTrabajo & strDll) Then
540         MsgBox "No existe: " & strCarpetaTrabajo & strDll
550         lst1.Clear
560         Exit Sub
570      End If
580      If Not fs.FolderExists(strOutdir) Then
590         MsgBox "No existe: " & strOutdir
600         lst1.Clear
610         Exit Sub
620      End If
         
630      Timer1.Enabled = True
         
640      Shell strEjecuta, vbHide
650      Do While fs.FileExists(strErrores) = False
660         DoEvents
670      Loop

680      PID = El_pid("vb6c.exe")

700      Do While FileLen(strErrores) = 0 And Timer1.Enabled = True
710         DoEvents
720      Loop

730      Timer1.Enabled = False

740      Set ts = fs.OpenTextFile(strErrores)
         
750      Do While Not ts.AtEndOfStream
            strError = ts.ReadLine
760         lst1.AddItem strError
            getlistboxHScrollBar strError
770      Loop
780      Set ts = Nothing
800      lst1.ListIndex = lst1.ListCount - 1
         
810      If PID > 0 Then
811         aEstadisticas(ix) = elapsed
820         Borrar_Sobras (ix)
            If Not (cmdCompatibilidad.Picture = ImageList.ListImages(6).Picture) Then 'Compatibilidad proyecto
830            Compactar (ix)
            End If
840      End If
860      Set fs = Nothing

861      ProgressBar.Value = 0.1
         
         'Cambiar proyecto
         If cmdCompatibilidad.Picture = ImageList.ListImages(6).Picture Then
            SetCompProyOFF strCarpetaTrabajo, strDll
         End If
         
870      Exit Sub
         
ErrorHandler:
880      strAction = "Compilar " & Err.Description & Erl
890      Timer1.Enabled = False
900      lst1.AddItem strAction

End Sub

Private Sub HabilitarControles()
   cmdCompilar.Enabled = True
   Command2.Enabled = NombrePC = "PC3" Or NombrePC = "PC5"
   If NombrePC <> "PC3" And NombrePC <> "PC5" Then
      Command3.Enabled = True
      cmdRegistrar.Enabled = True
   End If
   Command5.Enabled = True
   cmdAutocheck.Enabled = True
   Frame1.Enabled = True
   Frame2.Enabled = True
   Frame3.Enabled = True
   Frame4.Enabled = True
   cmdSalvar.Enabled = True
   cmbCarpeta.Enabled = True
   If NombrePC = "ADRIAN" Or NombrePC = "JUANJO" Or (NombrePC = "PC3" And UCase(txtAvisar.Text) = "ADRIAN") Then
      cmdCompatibilidad.Enabled = True
   Else
      cmdCompatibilidad.Enabled = False
   End If
   Orden.Enabled = True
   Screen.MousePointer = vbDefault
   Form1.Caption = "Compilador - Algoritmo S.A."
   
   iTiempoTotal = 0

   If cmdStart.Picture = ImageList.ListImages(2).Picture Then
      DTPicker.Enabled = True
   End If

   cmdReferencias.Enabled = True
End Sub
Private Sub DeshabilitarControles()
   cmdCompilar.Enabled = False
   Command2.Enabled = False
   Command3.Enabled = False
   Command5.Enabled = False
   cmdAutocheck.Enabled = False
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   Frame4.Enabled = False
   cmdSalvar.Enabled = False
   cmbCarpeta.Enabled = False
   cmdRegistrar.Enabled = False
   cmdCompatibilidad.Enabled = False
   Screen.MousePointer = vbArrowHourglass
   Orden.Enabled = False
   DTPicker.Enabled = False
   StatusBar.Panels(2).Text = ""
   cmdReferencias.Enabled = False
End Sub

Private Sub cmbVersion_Click()
      If NombrePC <> "PC3" And NombrePC <> "PC5" Then
         txtActTesting.Text = Left(txtActTesting.Text, InStrRev(txtActTesting.Text, "1.0") - 1) & cmbVersion.Text
      End If
End Sub
Private Sub Command2_Click()
   
10       On Error GoTo ErrorHandler
   
         If chkCopiaCierra2.Value = vbUnchecked And Not bHayErrores Then
11          CopiarReadme
         End If

         If chkCopiaCierra2.Value = vbChecked And Not bHayErrores Then
         Else
20          If MsgBox("¿Esta seguro que desea copiar los archivos marcados a " & vbCrLf & txtActClientes.Text & "?", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
         End If
   
30       bTesting = False
   
40       For ix = 0 To CANTIDAD_TOTAL_DLL
50          If chkComponente(ix) = vbChecked Then
60             CopiarArchivos (ix)
70          End If
80       Next
90       lst1.AddItem "Copia Terminada"
100      lst1.ListIndex = lst1.ListCount - 1
         
110      Exit Sub
   
ErrorHandler:
120      strAction = "Command2_Click " & Err.Description & Erl
130      lst1.AddItem strAction
End Sub
Private Sub Command3_Click()

10       On Error GoTo ErrorHandler
         
         If Len(txtCarpeta.Text) = 0 Then Exit Sub
         
20       bTesting = True
   
30       For ix = 0 To CANTIDAD_TOTAL_DLL
40          If chkComponente(ix) = vbChecked Then
50             CopiarArchivos (ix)
60          End If
70       Next
80       lst1.AddItem "Copia Terminada"
90       lst1.ListIndex = lst1.ListCount - 1
         
100      Exit Sub
   
ErrorHandler:
110      strAction = "Command3_Click " & Err.Description & Erl
120      lst1.AddItem strAction
End Sub
Private Sub CopiarArchivos(ix As Integer)
   
10       On Error GoTo ErrorHandler
   
20       Set fs = CreateObject("Scripting.FileSystemObject")
   
30       If InStr(chkComponente(ix).Caption, "DS") Or _
            InStr(chkComponente(ix).Caption, "SP") Or _
            InStr(chkComponente(ix).Caption, "DataAccess") Or _
            InStr(chkComponente(ix).Caption, "DataShare") Then
70          strOutdir = GetOutdir & "\server\"
90       Else
            If InStr(chkComponente(ix).Caption, "AlgInterop") Or InStr(chkComponente(ix).Caption, "AlgMobile") Then
               strOutdir = GetOutdir & "\services\"
            Else
100            strOutdir = GetOutdir & "\"
            End If
110      End If
   
120      If bTesting And Len(txtCarpeta.Text) > 0 Then
130         If Not fs.FolderExists(txtActTesting.Text & "\" & txtCarpeta.Text) Then
140            fs.CreateFolder (txtActTesting.Text & "\" & txtCarpeta.Text)
150         End If
160      End If
   
170      lst1.AddItem "Copiando " & chkComponente(ix).Caption
171      lst1.ListIndex = lst1.ListCount - 1
         
180      If fs.FileExists(strOutdir & chkComponente(ix).Caption) Then
190            If InStr(chkComponente(ix).Caption, "DS") Or InStr(chkComponente(ix).Caption, "SP") Or InStr(chkComponente(ix).Caption, "DataAccess") Or InStr(chkComponente(ix).Caption, "DataShare") Then
200               If bTesting Then
210                  FileCopy strOutdir & chkComponente(ix).Caption, txtActTesting.Text & "\" & IIf(Len(txtCarpeta.Text) > 0, txtCarpeta.Text & "\", "") & chkComponente(ix).Caption
220               Else
230                  FileCopy strOutdir & chkComponente(ix).Caption, txtActClientes.Text & "\Server\" & chkComponente(ix).Caption
231                  FileCopy strOutdir & Replace(chkComponente(ix).Caption, ".dll", ".rar"), txtActClientes.Text & "\Server\" & Replace(chkComponente(ix).Caption, ".dll", ".rar")
240               End If
250            Else
                  If InStr(chkComponente(ix).Caption, "AlgInterop") Or InStr(chkComponente(ix).Caption, "AlgMobile") Then
                     If bTesting Then
                        FileCopy strOutdir & chkComponente(ix).Caption, txtActTesting.Text & "\" & IIf(Len(txtCarpeta.Text) > 0, txtCarpeta.Text & "\", "") & chkComponente(ix).Caption
                     Else
                        FileCopy strOutdir & chkComponente(ix).Caption, txtActClientes.Text & "\Services\" & chkComponente(ix).Caption
                        FileCopy strOutdir & Replace(chkComponente(ix).Caption, ".dll", ".rar"), txtActClientes.Text & "\Services\" & Replace(chkComponente(ix).Caption, ".dll", ".rar")
                     End If

                  Else
                  
260                  If bTesting Then
270                     FileCopy strOutdir & chkComponente(ix).Caption, txtActTesting.Text & "\" & IIf(Len(txtCarpeta.Text) > 0, txtCarpeta.Text & "\", "") & chkComponente(ix).Caption
280                  Else
290                     FileCopy strOutdir & chkComponente(ix).Caption, txtActClientes.Text & "\Client\" & chkComponente(ix).Caption
291                     FileCopy strOutdir & Replace(chkComponente(ix).Caption, ".dll", ".rar"), txtActClientes.Text & "\Client\" & Replace(chkComponente(ix).Caption, ".dll", ".rar")
300                  End If
                  End If
310            End If
320      End If
   
330      Set fs = Nothing
   
340      Exit Sub
   
ErrorHandler:
350      strAction = "CopiarArchivos " & Err.Description & " " & Err.Number & " " & Erl
360      lst1.AddItem strAction
End Sub
Private Sub Cargar_cmbCarpeta()
   
   strNombreUsuario = GetRegistryValue(HKEY_CURRENT_USER, "Environment", "SSUSER")
   If InStr(strNombreUsuario, "comp") > 0 Then
      cmbCarpeta.AddItem "C:\VSS Carpetas de Trabajo Comp"
   Else
      cmbCarpeta.AddItem "C:\VSS Carpetas de Trabajo"
   End If
   cmbCarpeta.AddItem "\\alg01\d\Software"
   cmbCarpeta.AddItem "\\alg01\d\Otros\Versiones Fuentes\" & cmbVersion.List(1)

   Select Case NombrePC
      Case "PC3"
         cmbCarpeta.ListIndex = 1

      Case "PC5"
         cmbCarpeta.ListIndex = 2
         cmbVersion.ListIndex = 1
      Case Else
         cmbCarpeta.ListIndex = 0
   End Select
   
End Sub
Private Sub Borrar_Sobras(ix As Integer)
   
   On Error Resume Next
   
   If fs.FileExists(strOutdir & "\" & Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & ".lib") Then
      fs.DeleteFile (strOutdir & "\" & Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & ".lib")
   End If
   If fs.FileExists(strOutdir & "\" & Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & ".exp") Then
      fs.DeleteFile (strOutdir & "\" & Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & ".exp")
   End If
   'If fs.FileExists(strErrores) Then fs.DeleteFile (strErrores)
End Sub

Private Sub Command4_Click()

   Shell "explorer.exe " & txtActTesting.Text & "\" & IIf(Len(txtCarpeta.Text) > 0, txtCarpeta.Text & "\", ""), vbMaximizedFocus
   
End Sub

Private Sub Command8_Click()
   
   Shell "explorer.exe " & txtActClientes.Text, vbMaximizedFocus
   
End Sub
Private Sub Command5_Click()
      Dim strDll       As String
      Dim strProyectos As String
      Dim strGrupo     As String
      Dim fNumber1     As Integer

10       On Error GoTo ErrorHandler
   
20       strProyectos = ""
30       For ix = CANTIDAD_TOTAL_DLL To 0 Step -1
40          If chkComponente(ix) = vbChecked Then

50             strCarpetaTrabajo = GetCarpetaTrabajo(chkComponente(ix).Caption)
               
180            strDll = Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & "\" & Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & ".vbp"
               
               If chkComponente(ix).Caption = "AlgInterop.dll" Then
                  strProyectos = strProyectos & _
                                 "StartupProject=" & strCarpetaTrabajo & strDll & vbCrLf
               Else
190               strProyectos = strProyectos & _
                                 "Project=" & strCarpetaTrabajo & strDll & vbCrLf
               End If
200         End If
210      Next
         
220      strGrupo = "VBGROUP 5.0" & vbCrLf & strProyectos & _
                    IIf(InStr(strProyectos, "AlgInterop") > 0, "", "StartupProject=c:\VSS Carpetas de Trabajo\Inicio\Inicio.vbp")
              
230      fNumber1 = FreeFile
240      Open "C:\WINDOWS\Temp\Compilador.vbg" For Output As fNumber1
250      Print #fNumber1, strGrupo
260      Close #fNumber1
   
270      Set fs = CreateObject("Scripting.FileSystemObject")
280      If fs.FileExists(EJECUTABLE_VB6_X86) Then
290         Shell EJECUTABLE_VB6_X86 & " C:\WINDOWS\Temp\Compilador.vbg", vbMaximizedFocus
300      Else
310         Shell EJECUTABLE_VB6_X64 & " C:\WINDOWS\Temp\Compilador.vbg", vbMaximizedFocus
320      End If
330      Set fs = Nothing
   
340      Exit Sub
   
ErrorHandler:
350      strAction = "Command5_Click " & Err.Description & " " & Err.Number & " " & Erl
360      lst1.AddItem strAction
End Sub
Private Sub Command6_Click()
   
   Dim strSDP As String

   Set fs = CreateObject("Scripting.FileSystemObject")
   
   strSDP = "http://sdp.algoritmo.com.ar:4567/WorkOrder.do?woMode=viewWO&woID=" & txtCarpeta.Text & "&"
   
   If fs.FileExists("C:\Archivos de programa\Mozilla Firefox\firefox.exe") Then
      Shell "C:\Archivos de programa\Mozilla Firefox\firefox.exe " & strSDP, vbMaximizedFocus
   Else
      Shell "C:\Archivos de programa\Internet Explorer\iexplore.exe " & strSDP, vbMaximizedFocus
   End If
   
   Set fs = Nothing
   
End Sub
Private Sub cmdCompatibilidad_Click()
   
   If cmdCompatibilidad.Picture = ImageList.ListImages(6).Picture Then
'      strCompatibilidad = vbNullString
      cmdCompatibilidad.Picture = ImageList.ListImages(5).Picture
   Else
'      strCompatibilidad = " /d CompatibleMode=" & Chr(34) & "1" & Chr(34) & " VersionCompatible32=" & Chr(34) & "" & Chr(34) 'Compatibilidad de Proyecto
      cmdCompatibilidad.Picture = ImageList.ListImages(6).Picture
   End If
   
End Sub
Private Sub cmdStart_Click()
   If cmdStart.Picture = ImageList.ListImages(7).Picture Then
      cmdStart.Picture = ImageList.ListImages(2).Picture
      DTPicker.Enabled = True
   Else
      cmdStart.Picture = ImageList.ListImages(7).Picture
      DTPicker.Enabled = False
   End If
End Sub
Private Sub Avisar(strLista As String)
   Dim strEjecuta      As String
   Dim iRespuesta      As Integer
   Dim aSplit()        As String
   
   On Error GoTo ErrorHandler
   
   aSplit() = Split(Trim(txtAvisar.Text), ";")
   For ix = LBound(aSplit) To UBound(aSplit)
      If bCompCancelada Or bCompCerrado Then
'         If W10 Then
            strEjecuta = "cmd.exe /K " & Chr(34) & _
                         "msg /server:" & Trim(aSplit(ix)) & " console " & "Se CANCELO la compilación !!!" & Chr(34) & " & exit"
'         Else
'                  strEjecuta = "cmd.exe /K " & Chr(34) & _
'                      "net send " & Trim(aSplit(ix)) & " " & "Se CANCELO la compilación !!!" & Chr(34) & " & exit"
'
'         End If
      Else
         strEjecuta = "cmd.exe /K " & Chr(34) & _
                      "msg /server:" & Trim(aSplit(ix)) & " console " & "Se compilaron los siguientes proyectos: " & strLista & IIf(bHayErrores = True, " CON ERRORES !!!", "") & IIf(Len(txtCarpeta.Text) > 0, "(TP: " & txtActTesting.Text & "\" & txtCarpeta.Text & ")", "") & Chr(34) & " & exit"
      End If

      iRespuesta = Shell(strEjecuta, vbHide)
      
      lst1.AddItem "Mensaje a " & Trim(aSplit(ix)) & " enviado " & CStr(Time)
      lst1.ListIndex = lst1.ListCount - 1
      getlistboxHScrollBar "Mensaje a " & Trim(aSplit(ix)) & " enviado " & CStr(Time)
   Next ix
   
'   strEjecuta = "cmd.exe /K " & Chr(34) & _
'                "msg " & Trim(txtAvisar.Text) & " " & "Se compilaron los siguientes proyectos: " & strLista & Chr(34) & " & exit"

   Exit Sub
   
ErrorHandler:
   strAction = "Avisar " & Err.Description & Erl
   lst1.AddItem strAction

End Sub
Private Sub Compactar(ix As Integer)
   Dim strEjecuta As String
   
   On Error Resume Next
   
   bTerminoRAR = False
   
   If InStr(chkComponente(ix).Caption, "DS") Or _
      InStr(chkComponente(ix).Caption, "SP") Or _
      InStr(chkComponente(ix).Caption, "DataAccess") Or _
      InStr(chkComponente(ix).Caption, "DataShare") Then
      strOutdir = GetOutdir & "\server\"
   Else
      If InStr(chkComponente(ix).Caption, "AlgInterop") Or InStr(chkComponente(ix).Caption, "AlgMobile") Then
         strOutdir = GetOutdir & "\services\"
      Else
         strOutdir = GetOutdir & "\"
      End If
   End If
   
   strEjecuta = "cmd.exe /K " & Chr(34) & Chr(34) & _
                 "C:\Archivos de programa\WinRAR\rar.exe" & Chr(34) & " -ep a " & Chr(34) & strOutdir & Replace(chkComponente(ix).Caption, ".dll", "") & Chr(34) & " " & Chr(34) & strOutdir & chkComponente(ix).Caption & Chr(34) & Chr(34) & " & exit"

'  strEjecuta = "cmd.exe /c start " & Chr(34) & "Compactando..." & Chr(34) & " /belownormal " & Chr(34) & _
'                "C:\Archivos de programa\WinRAR\rar.exe" & Chr(34) & " a " & Chr(34) & strOutdir & Replace(chkComponente(ix).Caption, ".dll", "") & Chr(34) & " " & Chr(34) & strOutdir & chkComponente(ix).Caption & Chr(34)

   lngID_RAR = Shell(strEjecuta, vbHide) 'obtiene el pid directamente a la variable lngID_RAR
   
   Do While bTerminoRAR = False
      DoEvents
   Loop
   
   If NombrePC = "PC3" Or NombrePC = "PC5" Then Sleep 5000
   
   DoEvents
   
   lst1.AddItem "Dll compactada"
   lst1.AddItem ""
   lst1.AddItem "________________________________________________________________________"
   'lst1.AddItem "   (¯`·._.·(¯`·._.·(¯`·._.·(¯`·._.·(¯`·._.·(¯`·._.·(¯`·._.·(¯`·._.·(¯`·._.··._.·´¯)·._.·´¯)·._.·´¯)·._.·´¯)·._.·´¯)·._.·´¯)·._.·´¯)·._.·´¯)·._.·´¯)"
   lst1.AddItem ""
   lst1.ListIndex = lst1.ListCount - 1
   
End Sub

Private Sub Form_Terminate()
   End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Timer1.Enabled = True Then
      bCompCerrado = True
      Avisar ""
   End If
   cmdPararCompilacion_Click
   End
End Sub

Private Sub Timer1_Timer()
Dim strUltimo As String
Dim strTotal  As String

   elapsed = elapsed + 1
   
   If elapsed <= ProgressBar.Max And ProgressBar.Max > 0 Then
      ProgressBar.Value = elapsed
   End If
   
   If ProgressBar.Max < 60 Then
      strUltimo = "  (Ultima: " & ProgressBar.Max & " Segundos)"
   Else
      strUltimo = "  (Ultima: " & ProgressBar.Max \ 60 & " Minutos " & (ProgressBar.Max - (ProgressBar.Max \ 60) * 60) & " Segundos)"
   End If
   
   If elapsed < 60 Then
      lst1.List(lst1.ListCount - 1) = "Compilando " & chkComponente(ix).Caption & "... " & "Tiempo Usado: = " & elapsed & " Segundos " & strUltimo
   Else
      lst1.List(lst1.ListCount - 1) = "Compilando " & chkComponente(ix).Caption & "... " & "Tiempo Usado: = " & elapsed \ 60 & " Minutos " & (elapsed - (elapsed \ 60) * 60) & " Segundos " & strUltimo
   End If
   Me.Caption = Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & " " & Int(elapsed * 100 / ProgressBar.Max) & "%"
   
   'Calculo del tiempo total restante en el caption
   If iTiempoTotal = 0 And elapsed = 1 Then
      iTiempoTotal = iTiempo + iTotDLLS
   End If
   iTiempoTotal = iTiempoTotal - 1
   If iTiempoTotal < 0 Then
      iTiempoTotal = 0
   End If
   If iTiempoTotal < 60 Then
      strTotal = " 00:" & Format(iTiempoTotal, "00")
   Else
      strTotal = " " & Format(iTiempoTotal \ 60, "00") & ":" & Format((iTiempoTotal - (iTiempoTotal \ 60) * 60), "00")
   End If
   Me.Caption = Me.Caption & strTotal
   'Calculo del tiempo total restante en el caption
   
   ProgressBar.ToolTipText = Me.Caption
End Sub
Private Sub cmdSalvar_Click()
Dim fNumber1     As Integer
   
   fNumber1 = FreeFile
   strOutdir = GetOutdir
   Open strOutdir & "\Compilador.dat" For Output As fNumber1
   
   For ix = 0 To CANTIDAD_TOTAL_DLL
      If chkComponente(ix) = vbChecked Then
         Print #fNumber1, Trim(ix)
      End If
   Next

   Close #fNumber1

End Sub
Private Sub Revisar_checks()
Dim strContents As String
Dim objReadFile As Object

   Set fs = CreateObject("Scripting.FileSystemObject")
   
   strOutdir = GetOutdir
   
   If fs.FileExists(strOutdir & "\Compilador.dat") Then
      Set objReadFile = fs.OpenTextFile(strOutdir & "\Compilador.dat", 1)
      While Not objReadFile.AtEndOfStream
         ix = objReadFile.ReadLine
         chkComponente(ix).Value = vbChecked
      Wend
      objReadFile.Close
      Set objReadFile = Nothing
   End If
   Set fs = Nothing
   
End Sub
Private Sub Enviar_Mail()
Dim strSendTo   As String
Dim strMensaje  As String
Dim aSplit()    As String

   Dim Nombre As String, ret As Long, strUsuario As String
   
   On Error Resume Next
   
   If Not (NombrePC = "ADRIAN" Or (NombrePC = "PC3" And UCase(txtAvisar.Text = "ADRIAN"))) Then Exit Sub

   Nombre = Space$(250)
   ret = Len(Nombre)
   If GetUserName(Nombre, ret) = 0 Then
      strUsuario = vbNullString
   Else
      strUsuario = Left$(Nombre, ret - 1)
   End If
   
   If NombrePC = "ADRIAN" Or (NombrePC = "PC3" And UCase(txtAvisar.Text = "ADRIAN")) Then
      strSendTo = "adrian@algoritmo.com.ar"
   End If
   For ix = 1 To lst1.ListCount
      strMensaje = strMensaje & lst1.List(ix) & vbCrLf
   Next
   If Len(txtCarpeta.Text) > 0 Then
      strMensaje = strMensaje & "Incidente: " & txtActTesting.Text & "\" & txtCarpeta.Text
   End If

   EnviarEMail "Compilación de " & strUsuario, strMensaje, strSendTo, "Compilador"

End Sub
Private Sub cmdRegistrar_Click()
   For ix = 0 To CANTIDAD_TOTAL_DLL
      If chkComponente(ix) = vbChecked Then
         Desregistrar_Cliente chkComponente(ix).Caption
         Registrar_Cliente chkComponente(ix).Caption, "ALG01"
      End If
   Next
End Sub
Private Sub Desregistrar_Cliente(strDll As String)
   If InStr(chkComponente(ix).Caption, "DS") Or _
      InStr(chkComponente(ix).Caption, "SP") Or _
      InStr(chkComponente(ix).Caption, "DataAccess") Or _
      InStr(chkComponente(ix).Caption, "DataShare") Then
      strOutdir = GetOutdir & "\server\"
   Else
      If InStr(chkComponente(ix).Caption, "AlgInterop") Or InStr(chkComponente(ix).Caption, "AlgMobile") Then
         strOutdir = GetOutdir & "\services\"
      Else
         strOutdir = GetOutdir & "\"
      End If
   End If
   lst1.AddItem "Desregistrando Dll..."
   lst1.ListIndex = lst1.ListCount - 1
   
   WaitForShelledApp "regsvr32 /u /s " & """" & strOutdir & strDll & """"
End Sub

Private Sub Registrar_Cliente(strDll As String, strPone As String)
   lst1.AddItem "Registrando Dll a ALG01..."
   lst1.ListIndex = lst1.ListCount - 1
   
   WaitForShelledApp "regsvr32 /s " & """" & "\\" & strPone & "\D\Algoritmo\Componentes Client\" & strDll & """"
   lst1.AddItem "Proceso Terminado"
   lst1.ListIndex = lst1.ListCount - 1
End Sub
Private Sub WaitForShelledApp(strproceso As String)
Dim ProcessId As Long
Dim hProcess As Long
Dim ExitCode As Long

   ProcessId = Shell(strproceso, vbHide)
   hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId)
   Do
      DoEvents
      Call GetExitCodeProcess(hProcess, ExitCode)
   Loop While ExitCode > 0
   CloseHandle hProcess
End Sub
Public Sub EnviarEMail(ByVal strAsunto As String, _
                        ByVal strMensaje As String, _
                        ByVal strPara As String, _
                        ByVal strRemitente As String)

Dim objMessage As Object
Dim aAttach()  As String
Dim i As Integer
   
   
   On Error Resume Next

   '************************************************************************
   '********    ENVIO DE MAIL USANDO SMTP POR CDO                   ********
   '************************************************************************
   
   ' Ver: http://www.paulsadowski.com/WSH/cdo.htm
   '
   Const cdoSendUsingPort = 2 'Enviar Mensaje Por la Red Usando SMTP.
   
   'Const cdoAnonymous = 0 'Do not authenticate
   Const cdoBasic = 1 'basic (clear-text) authentication
   'Const cdoNTLM = 2 'NTLM

   Set objMessage = CreateObject("CDO.Message")
   objMessage.Subject = strAsunto
   objMessage.From = strRemitente
   objMessage.To = strPara
   objMessage.TextBody = Replace(strMensaje, "|", vbCrLf)
   
   '==This section provides the configuration information for the remote SMTP server.
   
   objMessage.Configuration.Fields.Item _
   ("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
   
   'Name or IP of Remote SMTP Server
   objMessage.Configuration.Fields.Item _
   ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "ALG02"
               
'   If mvarDefault.Autenticar = Si Then
'      'Type of authentication, NONE, Basic (Base64 encoded), NTLM
'      objMessage.Configuration.Fields.Item _
'      ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
'
'      'Your UserID on the SMTP server
'      objMessage.Configuration.Fields.Item _
'      ("http://schemas.microsoft.com/cdo/configuration/sendusername") = mvarDefault.Usuario
'
'      'Your password on the SMTP server
'      objMessage.Configuration.Fields.Item _
'      ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mvarDefault.Password
'
'   End If
   'Server port (typically 25)
   objMessage.Configuration.Fields.Item _
   ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587 '25
   
   'Use SSL for the connection (False or True)
   objMessage.Configuration.Fields.Item _
   ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
   
   'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
   objMessage.Configuration.Fields.Item _
   ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
   
   objMessage.Configuration.Fields.Update
   
   '==End remote SMTP server configuration section==
   
   objMessage.Send

   Set objMessage = Nothing

GestErr:


End Sub

Function El_pid(proceso As String) As Long
    
    Dim hSnapShot As Long, uProcess As PROCESSENTRY32
    hSnapShot = CreateToolhelp32Snapshot(&H1 Or &H2 Or &H4 Or &H8, 0&)
    uProcess.dwSize = Len(uProcess)
    Dim r As Long: r = Process32First(hSnapShot, uProcess)
    
    Do While r
        If LCase(Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))) = LCase(proceso) Then
          El_pid = uProcess.th32ProcessID
          Exit Do
        End If
        r = Process32Next(hSnapShot, uProcess)
    Loop
    CloseHandle hSnapShot

End Function

Private Sub TimerStart_Timer()
   If Not DTPicker.Enabled Then
      StatusBar.Panels(2).Text = ""
      Exit Sub
   End If
   
   If TimerStart.Enabled And Format(Time, "hh:mm:ss AM/PM") = Format(DTPicker.Value, "hh:mm:ss AM/PM") Then
      TimerStart.Enabled = False
      If NombrePC = "PC3" Or NombrePC = "PC5" Then
         chkCopiaCierra2.Value = vbChecked
         chkCopiaCierra2.Enabled = False
      End If
      cmdCompilar_Click
      TimerStart.Enabled = True
   End If
   If Format(DTPicker.Value, "hh:mm:ss") <> "00:00:00" And Format(DTPicker.Value, "hh:mm:ss") > Format(Time, "hh:mm:ss") Then
      StatusBar.Panels(2).Text = Format(TimeValue(Format(Time, "hh:mm:ss")) - TimeValue(Format(DTPicker.Value, "hh:mm:ss")), "hh:mm:ss")
   Else
      StatusBar.Panels(2).Text = ""
   End If

End Sub

Private Sub txtCarpeta_Change()
   Clipboard.Clear
   If Not bBusqueSS Then strSS = GetSourceSafe
   Sleep 100
   Clipboard.SetText strSS & txtActTesting.Text & "\" & IIf(Len(txtCarpeta.Text) > 0, txtCarpeta.Text, "")
End Sub

Private Function BuscarCompatibilidad(ix As Integer) As String
      Dim strContents As String
      Dim objReadFile As Object
      Dim strDll      As String
      Dim strLinea    As String
         
10       On Error GoTo ErrorHandler
         
20       Set fs = CreateObject("Scripting.FileSystemObject")
         
50       strCarpetaTrabajo = GetCarpetaTrabajo(chkComponente(ix).Caption)

240      strDll = Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & "\" & Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & ".vbp"
250      If fs.FileExists(strCarpetaTrabajo & strDll) Then
260         Set objReadFile = fs.OpenTextFile(strCarpetaTrabajo & strDll, 1)
270         While Not objReadFile.AtEndOfStream
280            strLinea = objReadFile.ReadLine
290            If InStr(strLinea, "CompatibleEXE32") Then
300               BuscarCompatibilidad = Replace(strLinea, "CompatibleEXE32=", "")
310            End If
320         Wend
330         objReadFile.Close
340         Set objReadFile = Nothing
350      End If
360      Set fs = Nothing
         
370      Exit Function
         
ErrorHandler:
380      strAction = "BuscarCompatibilidad " & Err.Description & Erl
390      lst1.AddItem strAction
End Function
Private Function NombrePC() As String
Dim dwLen As Long
Dim strString As String
   dwLen = MAX_COMPUTERNAME_LENGTH + 1
   strString = String(dwLen, "X")
   GetComputerName strString, dwLen
   NombrePC = Left(strString, dwLen)
End Function

Private Sub Command10_Click()
   Shell "explorer.exe " & GetOutdir, vbMaximizedFocus
End Sub

Private Sub LeerEstadisticas()
      Dim strContents As String
      Dim objReadFile As Object
         
10       On Error GoTo ErrorHandler
         
20       ProgressBar.Value = 0.1
         
30       Set fs = CreateObject("Scripting.FileSystemObject")
         
40       strOutdir = GetOutdir
50       ix = 0
         
60       If fs.FileExists(strOutdir & "\Estadisticas.dat") Then
70          Set objReadFile = fs.OpenTextFile(strOutdir & "\Estadisticas.dat", 1)
80          While Not objReadFile.AtEndOfStream
90             aEstadisticas(ix) = objReadFile.ReadLine
100            ix = ix + 1
110         Wend
120         objReadFile.Close
130         Set objReadFile = Nothing
140      Else
150         For iz = 0 To CANTIDAD_TOTAL_DLL
160            aEstadisticas(iz) = 1
170         Next
180      End If
190      Set fs = Nothing
         
200      Exit Sub
               
ErrorHandler:
210      strAction = "LeerEstadisticas " & Err.Description & Erl
220      lst1.AddItem strAction
End Sub

Private Sub SalvarEstadisticas()
      Dim fNumber1     As Integer
         
10       On Error GoTo ErrorHandler
         
20       fNumber1 = FreeFile
30       strOutdir = GetOutdir
40       Open strOutdir & "\Estadisticas.dat" For Output As fNumber1
         
50       For iz = 0 To CANTIDAD_TOTAL_DLL
60           Print #fNumber1, Trim(aEstadisticas(iz))
70       Next

80       Close #fNumber1
         
90       Exit Sub
               
ErrorHandler:
100      strAction = "SalvarEstadisticas " & Err.Description & Erl
110      lst1.AddItem strAction
End Sub
Public Sub SetBackColor(ProgressBarHwnd As Long, RGBValue As Long)
    Call SendMessage(ProgressBarHwnd, SB_SETBKCOLOR, 0, _
      ByVal RGBValue)
End Sub
 
Public Sub SetBarColor(ProgressBarHwnd As Long, RGBValue As Long)
    Call SendMessage(ProgressBarHwnd, PBM_SETBARCOLOR, 0, _
        ByVal RGBValue)
End Sub

Private Function GetSourceSafe() As String
      Dim strEjecuta      As String
      Dim iRespuesta      As Integer
      Dim Nombre As String, ret As Long, strUsuario As String
      Dim objReadFile     As Object
      Dim strLinea        As String
      Dim ix              As Integer
      Dim fs              As Object
      Dim strOutdir       As String
            
10       On Error GoTo ErrorHandler
         
20       If NombrePC = "PC3" Or NombrePC = "PC5" Then Exit Function
         
30       Set fs = CreateObject("Scripting.FileSystemObject")
         
         strOutdir = GetOutdir
         
40       If fs.FileExists(strOutdir & "\vssfind.exe") = False Then Exit Function
         
50       GetSourceSafe = "*** " & Now & " ***" & vbCrLf
         If InStr(strNombreUsuario, "comp") > 0 Then
            GetSourceSafe = GetSourceSafe & "Usuario " & strNombreUsuario & vbCrLf
         End If
         
60       If fs.FileExists(strOutdir & "\SS.txt") Then
70          Set objReadFile = fs.OpenTextFile(strOutdir & "\SS.txt", 1)
80          While Not objReadFile.AtEndOfStream
90             strLinea = objReadFile.ReadLine
100            For ix = 0 To CANTIDAD_TOTAL_DLL
110               If chkComponente(ix) = vbChecked Then
120                  If InStr(UCase(strLinea), "/" & UCase(Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1)) & "/") > 0 Or _
                        InStr(UCase(strLinea), "/" & UCase("Shared") & "/") > 0 Then 'Agrego el shared a pedido de Matias
130                     strLinea = Left(strLinea, InStr(strLinea, ",") - 1)
140                     GetSourceSafe = GetSourceSafe & strLinea & vbCrLf
                        bBusqueSS = True
150                  End If
160               End If
170            Next
180         Wend
190         objReadFile.Close
200         Set objReadFile = Nothing
210      End If
220      Set fs = Nothing

230      GetSourceSafe = GetSourceSafe & vbCrLf
         
240      Exit Function
                     
ErrorHandler:
250      strAction = "GetSourceSafe " & Err.Description & Erl
260      lst1.AddItem strAction
End Function
Private Sub TiempoNecesario(ByVal ix As Integer, ByVal bSuma As Boolean)
         Dim ij       As Integer
         Dim bChecked As Boolean
         
10       On Error GoTo ErrorHandler

20       If bSuma Then
30          iTiempo = iTiempo + IIf(aEstadisticas(ix) = 0, 1, aEstadisticas(ix))
40       Else
50          iTiempo = iTiempo - IIf(aEstadisticas(ix) = 0, 1, aEstadisticas(ix))
60       End If
70       If iTiempo < 60 Then
80          Me.Caption = "Tiempo de Compilación Necesario: " & iTiempo & " Segundos "
90       Else
100         Me.Caption = "Tiempo de Compilación Necesario: " & iTiempo \ 60 & " Minutos " & (iTiempo - (iTiempo \ 60) * 60) & " Segundos "
110      End If

120      bChecked = False
         iTotDLLS = 0
130      For ij = 0 To CANTIDAD_TOTAL_DLL
140         If chkComponente(ij) = vbChecked Then
150             bChecked = True
                iTotDLLS = iTotDLLS + 1
160         End If
170      Next
180      If bChecked = False Then iTiempo = 0
         
190      Exit Sub
                     
ErrorHandler:
200      strAction = "TiempoNecesario " & Err.Description & Erl
210      lst1.AddItem strAction
End Sub

Private Sub SetSourceSafe()
      Dim strEjecuta      As String
      Dim iRespuesta      As Integer
      Dim Nombre As String, ret As Long, strUsuario As String
      Dim objReadFile     As Object
      Dim strLinea        As String
      Dim ix              As Integer
      Dim fs              As Object
      Dim PID As Long
               
10       On Error GoTo ErrorHandler
         
20       If NombrePC = "PC3" Or NombrePC = "PC5" Then Exit Sub
         
30       Set fs = CreateObject("Scripting.FileSystemObject")
         
40       If fs.FileExists(strOutdir & "\vssfind.exe") = False Then
            txtCarpeta.Visible = True
            Set fs = Nothing
            Exit Sub
         Else
            txtCarpeta.Visible = False
         End If
         
'50       Nombre = Space$(250)
'60       ret = Len(Nombre)
'70       If GetUserName(Nombre, ret) = 0 Then
'80          strUsuario = vbNullString
'90       Else
'100         strUsuario = Left$(Nombre, ret - 1)
'110      End If

         'D:\>vssfind --db \\alg01\d\software\vss\srcsafe.ini --recurse --type status --checked-out-to juanjo.sortino --delim-output - > lista.txt
         
120       strEjecuta = "cmd.exe /K " & Chr(34) & _
                            Chr(34) & strOutdir & "\vssfind" & Chr(34) & " --db \\alg01\d\software\vss\srcsafe.ini --recurse --type status --checked-out-to " & _
                            strNombreUsuario & " --delim-output - > " & Chr(34) & strOutdir & "\SS.txt" & Chr(34) & _
                            Chr(34) & " & exit"
               
130      PID = Shell(strEjecuta, vbHide)
         If PID > 0 Then
            SetPriority PID, RealTime
         End If

140      Exit Sub
                     
ErrorHandler:
150      strAction = "SetSourceSafe " & Err.Description & Erl
160      lst1.AddItem strAction
End Sub

Private Sub CargarUsuario(bVnc As Boolean)
'http://allapi.mentalis.org/apilist/GetTcpTable.shtml#
Dim TcpRow As MIB_TCPROW
Dim buff() As Byte
Dim lngRequired As Long
Dim lngStrucSize As Long
Dim lngRows As Long
Dim lngCnt As Long
Dim strTmp As String
Dim lstLine As ListItem
Dim strHost As String

Call GetTcpTable(ByVal 0&, lngRequired, 1)

If lngRequired > 0 Then
    ReDim buff(0 To lngRequired - 1) As Byte
    If GetTcpTable(buff(0), lngRequired, 1) = ERROR_SUCCESS Then
        lngStrucSize = LenB(TcpRow)
        CopyMemory lngRows, buff(0), 4
       
        For lngCnt = 1 To lngRows
            CopyMemory TcpRow, buff(4 + (lngCnt - 1) * lngStrucSize), lngStrucSize
            If ntohs(TcpRow.dwLocalPort) = 5900 Then 'puerto del vnc
               If bVnc Then
                  If TcpRow.dwState = MIB_TCP_STATE_ESTAB Then
                     strHost = GetInetAddrStr(TcpRow.dwRemoteAddr)
                     List.AddItem "(" & Replace(Right(GetInetAddrStr(TcpRow.dwRemoteAddr), 3), ".", "") & ") " & UCase(Recuperar_Nombre_Host(GetInetAddrStr(TcpRow.dwRemoteAddr)))
                     List.ListIndex = 0
                  End If
               Else
                  txtAvisar.Text = Recuperar_Nombre_Host(GetInetAddrStr(TcpRow.dwRemoteAddr))
               End If
            Else
               If Len(txtAvisar.Text) = 0 Then
                  txtAvisar.Text = Recuperar_Nombre_Host(GetInetAddrStr(TcpRow.dwRemoteAddr))
               End If
            End If
        Next
    End If
End If
If List.ListCount = 1 And bVnc And Timer1.Enabled = False Then
   txtAvisar.Text = Recuperar_Nombre_Host(strHost)
End If
End Sub
Private Function GetString(ByVal lpszA As Long) As String
    GetString = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal GetString, ByVal lpszA)
End Function
Private Function GetInetAddrStr(Address As Long) As String
    GetInetAddrStr = GetString(inet_ntoa(Address))
End Function

Private Function Recuperar_Nombre_Host(ByVal direccion_IP As String) As String
         Dim PH As Long, hDir As Long, nb As Long
         Dim W As WSADATA
         
10       On Error GoTo ErrorHandler
         
20       If WSAStartup(&H101, W) = 0 Then
30          hDir = inet_addr(direccion_IP)
              
            'Si devuelve -1 dió error
40          If hDir <> -1 Then
              
50             PH = gethostbyaddr(hDir, 4, 2)
60             If PH <> 0 Then
70                CopyMemory PH, ByVal PH, 4
80                nb = lstrlen(ByVal PH)
90                If nb > 0 Then
100                  direccion_IP = Space$(nb)
110                  CopyMemory ByVal direccion_IP, ByVal PH, nb
120                  Recuperar_Nombre_Host = Replace(direccion_IP, ".algoritmo.local", "")
130               End If
140            Else
150                Recuperar_Nombre_Host = direccion_IP
160            End If
170            If WSACleanup() <> 0 Then
180               Recuperar_Nombre_Host = direccion_IP
190            End If
200         Else
210            Recuperar_Nombre_Host = direccion_IP
220         End If
230      Else
240         Recuperar_Nombre_Host = direccion_IP
250      End If
         
260      Exit Function
                           
ErrorHandler:
270      strAction = "recuperar_Nombre_Host " & Err.Description & Erl
280      lst1.AddItem strAction
End Function
Private Sub Timer_Timer()
   List.Clear
   'CargarUsuario True
   
   'Espero que termine vssfind.exe
   MonitorearVssfind
   MonitorearRAR
   DoEvents
   
'   If bnewells Then
'      'SetBarColor ProgressBar.hWnd, &HFF&
'      SetBarColor ProgressBar.hWnd, &HFF0000
'      bnewells = False
'   Else
'      'SetBarColor ProgressBar.hWnd, &H0&
'      SetBarColor ProgressBar.hWnd, &HFFFFFF
'      bnewells = True
'   End If
End Sub

Private Sub MonitorearVssfind()
Dim fs              As Object

   Set fs = CreateObject("Scripting.FileSystemObject")

   If fs.FileExists(strOutdir & "\vssfind.exe") = False Then Exit Sub
   
   'Espero que termine vssfind.exe
   lngID_Vss = El_pid("vssfind.exe")
   If lngID_Vss <> 0 Then
      txtCarpeta.Visible = False
      cmdAutocheck.Visible = False
   Else
      txtCarpeta.Visible = True
      cmdAutocheck.Visible = True
   End If
   
   Set fs = Nothing
   
End Sub
Private Sub MonitorearRAR()

   'Espero que termine rar.exe
   If lngID_RAR <> 0 Then
      bTerminoRAR = True
   Else
      bTerminoRAR = False
   End If
   
End Sub
Private Sub cmdAutocheck_Click()
Dim strEjecuta      As String
Dim iRespuesta      As Integer
Dim Nombre As String, ret As Long, strUsuario As String
Dim objReadFile     As Object
Dim strLinea        As String
Dim ix              As Integer
Dim fs              As Object
Dim strOutdir       As String
      
   On Error GoTo ErrorHandler
   
   If NombrePC = "PC3" Or NombrePC = "PC5" Then Exit Sub
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   
   strOutdir = GetOutdir
   
   If fs.FileExists(strOutdir & "\vssfind.exe") = False Then Exit Sub
   
   If fs.FileExists(strOutdir & "\SS.txt") Then
      Set objReadFile = fs.OpenTextFile(strOutdir & "\SS.txt", 1)
      While Not objReadFile.AtEndOfStream
         strLinea = objReadFile.ReadLine
         For ix = 0 To CANTIDAD_TOTAL_DLL
            If InStr(UCase(strLinea), "/" & UCase(Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1)) & "/") > 0 Then
               If CDate(Mid(strLinea, InStr(InStr(strLinea, ",") + 1, strLinea, ",") + 1, 10)) = Date Then
                  chkComponente(ix).Value = vbChecked
               End If
            End If
         Next
      Wend
      objReadFile.Close
      Set objReadFile = Nothing
   End If
   Set fs = Nothing

   Exit Sub
               
ErrorHandler:
   strAction = "cmdAutocheck_Click " & Err.Description & Erl
   lst1.AddItem strAction
End Sub

Private Sub CopiarReadme()
      Dim strfile    As String
      Dim strDestino As String
         
10       On Error GoTo ErrorHandler
         
20       Set fs = CreateObject("Scripting.FileSystemObject")
         
30       strDestino = "\\alg01\d\ftp\Descargas\SoftCereal\Readmes Versiones\" & Right(cmbVersion.Text, 3)
40       If Not (fs.FolderExists(strDestino)) Then
50          strAction = "La carpeta destino para readmes " & Right(cmbVersion.Text, 3) & " no existe " & strDestino
60          lst1.AddItem strAction
61          lst1.ListIndex = lst1.ListCount - 1
            getlistboxHScrollBar strAction
70          Exit Sub
80       End If
         
90       cdg1.InitDir = "\\stor1\shared\Readmes En Proceso"
100      cdg1.DialogTitle = "Mover readme a " & AddBackslash(strDestino)
110      cdg1.CancelError = True
120      cdg1.Filter = "Documentos de Word *.docx|*.docx|*.doc|*.doc|*.*|Todos los Archivos"
130      cdg1.FilterIndex = 1
140      cdg1.Flags = cdlOFNHideReadOnly + cdlOFNExtensionDifferent + cdlOFNOverwritePrompt
150      cdg1.DefaultExt = ".docx"
160      cdg1.ShowOpen
         
170      If cdg1.FileName <> "" Then
180         strfile = Replace(cdg1.FileName, "\\stor1\shared\Readmes En Proceso\", "")
190         FileCopy cdg1.FileName, AddBackslash(strDestino) & strfile
200         Kill cdg1.FileName
210         strAction = "Readme Movido a " & AddBackslash(strDestino) & strfile
220         lst1.AddItem strAction
221         lst1.ListIndex = lst1.ListCount - 1
            getlistboxHScrollBar strAction
230      End If
         
240      Set fs = Nothing
         
250      Exit Sub
               
ErrorHandler:
260      If Err.Number <> 32755 Then 'cancelar dialogo
270         strAction = "CopiarReadme " & Err.Description & Erl
280         lst1.AddItem strAction
290      End If
300      Set fs = Nothing
End Sub
  
Private Function Is_Transparent(ByVal hwnd As Long) As Boolean
'Función para saber si formulario ya es transparente.  Se le pasa el Hwnd del formulario en cuestión
On Error Resume Next
  
Dim Msg As Long
  
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
         
       If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
          Is_Transparent = True
       Else
          Is_Transparent = False
       End If
  
    If Err Then
       Is_Transparent = False
    End If
End Function
  
Private Function Aplicar_Transparencia(ByVal hwnd As Long, Valor As Integer) As Long
'Función que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Dim Msg As Long
  
   On Error Resume Next
     
   If Valor < 0 Or Valor > 255 Then
      Aplicar_Transparencia = 1
   Else
      Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
      Msg = Msg Or WS_EX_LAYERED
        
      SetWindowLong hwnd, GWL_EXSTYLE, Msg
        
      'Establece la transparencia
      SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA
     
      Aplicar_Transparencia = 0
     
   End If
     
     
   If Err Then
      Aplicar_Transparencia = 2
   End If
End Function

Function ServiceCommand(ByVal ServiceName As String, ByVal command As Long) As Boolean
' start/stop/pause/continue a service
' SERVICENAME is the
' COMMAND can be   0=Start, 1=Stop, 2=Pause, 3=Continue
'
' returns True if successful, False otherwise
' if any error, call Err.LastDLLError for more information
    Dim hSCM As Long
    Dim hService As Long
    Dim res As Long

    Dim query As Long
    Dim lpServiceStatus As SERVICE_STATUS

    ' first, check the command
    If command < 0 Or command > 3 Then Err.Raise 5

    ' open the connection to Service Control Manager, exit if error
    hSCM = OpenSCManager(vbNullString, vbNullString, GENERIC_EXECUTE)
    If hSCM = 0 Then Exit Function



    ' open the given service, exit if error
    hService = OpenService(hSCM, ServiceName, GENERIC_EXECUTE)
    If hService = 0 Then GoTo CleanUp

    'fetch the status
    query = QueryServiceStatus(hService, lpServiceStatus)

    ' start the service
    Select Case command
        Case 0
            ' to start a service you must use StartService
            res = StartService(hService, 0, 0)
        Case SERVICE_CONTROL_STOP, SERVICE_CONTROL_PAUSE, _
            SERVICE_CONTROL_CONTINUE
            ' these commands use ControlService API
            ' (pass a NULL pointer because no result is expected)
            res = ControlService(hService, command, lpServiceStatus)
    End Select
    If res = 0 Then GoTo CleanUp

    ' return success
    ServiceCommand = True

CleanUp:
        If hService Then CloseServiceHandle hService
        ' close the SCM
        CloseServiceHandle hSCM

End Function

Private Sub ActualizarLog()
      Dim fNumber1     As Integer
      Dim strCabecera  As String
      Dim strUsuario   As String

10       On Error GoTo ErrorHandler
   
20       fNumber1 = FreeFile
30       strOutdir = GetOutdir
40       Open strOutdir & "\Compilador.log" For Append As fNumber1
   
41       If Len(Trim(strNombreUsuario)) > 0 Then
42          strUsuario = strNombreUsuario
43       Else
44          strUsuario = txtAvisar.Text
45       End If
50       strCabecera = vbCrLf & "Usuario: " & strUsuario & " " & Now & " " & NombrePC
60       If NombrePC <> "PC3" And NombrePC <> "PC5" And Len(txtCarpeta.Text) > 0 Then
70          strCabecera = strCabecera & " SDP: " & txtCarpeta.Text
80       End If
   
90       Print #fNumber1, strCabecera
100      Print #fNumber1, cmbCarpeta.Text
110      For iz = 0 To CANTIDAD_TOTAL_DLL
120         If chkComponente(iz) = vbChecked Then
130            Print #fNumber1, chkComponente(iz).Caption
140         End If
150      Next
160      If chkCopiaCierra.Value = vbChecked Or chkCopiaCierra2.Value = vbChecked Then
170         Print #fNumber1, "Copiados Automaticamente"
180      End If
   
190      Close #fNumber1
   
200      Exit Sub
         
ErrorHandler:
210      strAction = "ActualizarLog " & Err.Description & Erl
220      lst1.AddItem strAction
End Sub

Private Sub SetCompProyON(ByVal strCarpetaTrabajo As String, ByVal strDll As String)

Dim fNumber1            As Integer
Dim strCabecera         As String
Dim strUsuario          As String
Dim strProyectoOriginal As String
Dim fs                  As Object
Dim strNombreProyecto   As String
Dim Attributes          As VbFileAttribute

   On Error GoTo ErrorHandler
   
   '-------------------------
   'Leo el proyecto Original y lo guardo como ".copia"
   Dim strContents As String
   Dim objReadFile As Object

   Set fs = CreateObject("Scripting.FileSystemObject")
   
   strNombreProyecto = strCarpetaTrabajo & strDll
   strProyectoOriginal = vbNullString
   If fs.FileExists(strNombreProyecto) Then
      Set objReadFile = fs.OpenTextFile(strNombreProyecto, 1)
      While Not objReadFile.AtEndOfStream
         strProyectoOriginal = strProyectoOriginal & objReadFile.ReadLine & vbCrLf
      Wend
      objReadFile.Close
      Set objReadFile = Nothing
   End If
   Set fs = Nothing
   
   FileCopy strNombreProyecto, strNombreProyecto & ".copia"
   
   '-------------------------
   'Creo el proyecto con Compatibilidad de proyecto
   strProyectoOriginal = Replace(strProyectoOriginal, "CompatibleMode=" & Chr(34) & "2" & Chr(34), "CompatibleMode=" & Chr(34) & "1" & Chr(34))
   strProyectoOriginal = Replace(strProyectoOriginal, "VersionCompatible32=" & Chr(34) & "1" & Chr(34), vbNullString)
   
   '      strCompatibilidad = " /d CompatibleMode=" & Chr(34) & "1" & Chr(34) & " VersionCompatible32=" & Chr(34) & "" & Chr(34) 'Compatibilidad de Proyecto
   
   Attributes = GetAttr(strNombreProyecto)
   If (Attributes And vbReadOnly) Then
      Attributes = Attributes - vbReadOnly
      SetAttr strNombreProyecto, Attributes
   End If
   Kill strNombreProyecto
   
   fNumber1 = FreeFile
   Open strNombreProyecto For Output As fNumber1
   Print #fNumber1, strProyectoOriginal
   Close #fNumber1
   
   lst1.AddItem "Compilando con Compatibilidad Proyecto"
   lst1.ListIndex = lst1.ListCount - 1
   lst1.AddItem ""
   lst1.ListIndex = lst1.ListCount - 1
   
   Exit Sub
   
ErrorHandler:
   strAction = "SetCompProyON " & Err.Description & Erl
   lst1.AddItem strAction
End Sub
Private Sub SetCompProyOFF(ByVal strCarpetaTrabajo As String, ByVal strDll As String)
Dim fs                  As Object
Dim strNombreProyecto   As String
Dim Attributes          As VbFileAttribute

   On Error GoTo ErrorHandler
   
   strNombreProyecto = strCarpetaTrabajo & strDll
   FileCopy strNombreProyecto & ".copia", strNombreProyecto

   Attributes = GetAttr(strNombreProyecto)
   Attributes = Attributes + vbReadOnly
   SetAttr strNombreProyecto, Attributes

   Kill strNombreProyecto & ".copia"
   
   Exit Sub
   
ErrorHandler:
   strAction = "SetCompProyOFF " & Err.Description & Erl
   lst1.AddItem strAction
End Sub
Private Sub Orden_Click()
Dim strMsg As String
   strMsg = ""
   For ix = 0 To CANTIDAD_TOTAL_DLL
      If chkComponente(ix) = vbChecked Then
         strMsg = strMsg & chkComponente(ix).Caption & vbCrLf
      End If
   Next
   If Len(strMsg) > 0 Then
      MsgBox strMsg, , "Orden de Compilación"
   End If
End Sub
Private Sub BuscarReferencias(ByVal iDllBuscada As Integer, ByVal ix As Integer)
Dim strContents As String
Dim objReadFile As Object
Dim strDll      As String
Dim strLinea    As String

   
   On Error GoTo ErrorHandler
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   
   strCarpetaTrabajo = GetCarpetaTrabajo(chkComponente(ix).Caption)
   
   strDll = Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & "\" & Left(chkComponente(ix).Caption, InStr(chkComponente(ix).Caption, ".") - 1) & ".vbp"
   If fs.FileExists(strCarpetaTrabajo & strDll) Then
      Set objReadFile = fs.OpenTextFile(strCarpetaTrabajo & strDll, 1)
      While Not objReadFile.AtEndOfStream
         strLinea = objReadFile.ReadLine
         If InStr(UCase(strLinea), UCase("Reference=")) > 0 And InStr(UCase(strLinea), "\" & UCase(chkComponente(iDllBuscada).Caption) & "#") > 0 Then
            If Len(strReferencias) = 0 Then
               lst1.AddItem "Referencias a " & chkComponente(iDllBuscada).Caption & ":"
            End If
            lst1.AddItem " ---> " & chkComponente(ix).Caption
            strReferencias = strReferencias & " " & chkComponente(ix).Caption
         End If
      Wend
      objReadFile.Close
      Set objReadFile = Nothing
   End If
   Set fs = Nothing

   Exit Sub
   
ErrorHandler:
   strAction = "BuscarReferencias " & Err.Description & Erl
   lst1.AddItem strAction
End Sub


Private Sub Command11_Click()
Dim iz  As Integer
   For iz = 10 To 17
      chkComponente(iz).Value = vbUnchecked
   Next
End Sub

Private Sub Command12_Click()
Dim iz  As Integer
   For iz = 10 To 17
      chkComponente(iz).Value = vbChecked
   Next
End Sub

Private Sub Command13_Click()
Dim iz  As Integer
   For iz = 0 To 9
      chkComponente(iz).Value = vbChecked
   Next
End Sub

Private Sub Command14_Click()
Dim iz  As Integer
   For iz = 0 To 9
      chkComponente(iz).Value = vbUnchecked
   Next
End Sub

Private Sub Command15_Click()
   'Varios
   chkComponente(18).Value = vbUnchecked
   chkComponente(19).Value = vbUnchecked
   chkComponente(32).Value = vbUnchecked
   chkComponente(33).Value = vbUnchecked
End Sub

Private Sub Command16_Click()
   'Varios
   chkComponente(18).Value = vbChecked
   chkComponente(19).Value = vbChecked
   chkComponente(32).Value = vbChecked
   chkComponente(33).Value = vbChecked
End Sub

Private Sub Command17_Click()
   For iz = 0 To CANTIDAD_TOTAL_DLL
      chkComponente(iz).Value = vbUnchecked
   Next
End Sub

Private Sub Command7_Click()
   chkComponente(20).Value = vbChecked
   chkComponente(21).Value = vbChecked
   chkComponente(22).Value = vbChecked
   chkComponente(23).Value = vbChecked
   chkComponente(24).Value = vbChecked
   chkComponente(25).Value = vbChecked
   chkComponente(26).Value = vbChecked
   chkComponente(27).Value = vbChecked
   chkComponente(28).Value = vbChecked
   chkComponente(29).Value = vbChecked
   chkComponente(30).Value = vbChecked
   chkComponente(31).Value = vbChecked
End Sub

Private Sub Command9_Click()
   chkComponente(23).Value = vbUnchecked
   chkComponente(22).Value = vbUnchecked
   chkComponente(28).Value = vbUnchecked
   chkComponente(24).Value = vbUnchecked
   chkComponente(26).Value = vbUnchecked
   chkComponente(20).Value = vbUnchecked
   chkComponente(21).Value = vbUnchecked
   chkComponente(27).Value = vbUnchecked
   chkComponente(25).Value = vbUnchecked
   chkComponente(29).Value = vbUnchecked
   chkComponente(30).Value = vbUnchecked
   chkComponente(31).Value = vbUnchecked
End Sub

Private Sub chkComponente_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Xmove = X
End Sub

Private Sub chkComponente_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If X > Xmove + 500 Then
      Select Case Index
         Case 23 'cereales
            If chkComponente(Index).Value = vbChecked Then
               chkComponente(16).Value = vbChecked
               chkComponente(7).Value = vbChecked
            Else
               chkComponente(16).Value = vbUnchecked
               chkComponente(7).Value = vbUnchecked
            End If
         Case 22 'gescom
            If chkComponente(Index).Value = vbChecked Then
               chkComponente(15).Value = vbChecked
               chkComponente(6).Value = vbChecked
            Else
               chkComponente(15).Value = vbUnchecked
               chkComponente(6).Value = vbUnchecked
            End If
         Case 28 'produccion
            If chkComponente(Index).Value = vbChecked Then
               chkComponente(17).Value = vbChecked
               chkComponente(8).Value = vbChecked
            Else
               chkComponente(17).Value = vbUnchecked
               chkComponente(8).Value = vbUnchecked
            End If
         Case 24 'fiscal
            If chkComponente(Index).Value = vbChecked Then
               chkComponente(14).Value = vbChecked
               chkComponente(5).Value = vbChecked
            Else
               chkComponente(14).Value = vbUnchecked
               chkComponente(5).Value = vbUnchecked
            End If
         Case 26 'contabilidad
            If chkComponente(Index).Value = vbChecked Then
               chkComponente(12).Value = vbChecked
               chkComponente(3).Value = vbChecked
            Else
               chkComponente(12).Value = vbUnchecked
               chkComponente(3).Value = vbUnchecked
            End If
         Case 20 'general
            If chkComponente(Index).Value = vbChecked Then
               chkComponente(13).Value = vbChecked
               chkComponente(4).Value = vbChecked
            Else
               chkComponente(13).Value = vbUnchecked
               chkComponente(4).Value = vbUnchecked
            End If
         Case 21 'general
            If chkComponente(Index).Value = vbChecked Then
               chkComponente(11).Value = vbChecked
               chkComponente(2).Value = vbChecked
            Else
               chkComponente(11).Value = vbUnchecked
               chkComponente(2).Value = vbUnchecked
            End If
      End Select
   End If
End Sub
Public Sub getlistboxHScrollBar(str As String)
   If Len(str) > lLengtext Then
      
      Select Case Len(str)
         Case Is <= 125
            'lLength = 600
            lLength = 500
         Case Is <= 145
            lLength = 700
         Case Else
            lLength = 1000
      End Select
      lLengtext = Len(str)
      Call SendMessage(lst1.hwnd, LB_SETHORIZONTALEXTENT, lLength, 0&)
   End If
End Sub

Private Function GetOutdir() As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FolderExists(OUTDIR) Then
        GetOutdir = OUTDIR
    Else
        GetOutdir = OUTDIR_X86
    End If
End Function

Private Function GetCarpetaTrabajo(ByVal strComponente As String) As String

   Dim strCarpetaTrabajo As String
    
   If InStr(strComponente, "DS") Or InStr(strComponente, "SP") Or InStr(strComponente, "BO") Or _
      InStr(strComponente, "AlgStdFunc") Or InStr(strComponente, "DataAccess") Or InStr(strComponente, "DataShare") Then
      Select Case cmbCarpeta.Text
         Case "C:\VSS Carpetas de Trabajo"
            strCarpetaTrabajo = "C:\VSS Carpetas de Trabajo\COM DLLs\"
         Case "\\alg01\d\Software"
            strCarpetaTrabajo = "\\alg01\d\Software\COM DLLs\"
         Case "\\alg01\d\Otros\Versiones Fuentes\" & cmbVersion.List(1)
            strCarpetaTrabajo = "\\alg01\d\Otros\Versiones Fuentes\" & cmbVersion.List(1) & "\COM DLLs\"
         Case Else
             strCarpetaTrabajo = cmbCarpeta.Text & "\COM DLLs\"
      End Select
   Else
      Select Case cmbCarpeta.Text
         Case "C:\VSS Carpetas de Trabajo"
            If InStr(strComponente, "ALGControls") Or InStr(strComponente, "PowerMask") Then
               strCarpetaTrabajo = "C:\VSS Carpetas de Trabajo\Controles\"
            Else
               If InStr(strComponente, "AlgInterop") Or InStr(strComponente, "AlgMobile") Then
                  strCarpetaTrabajo = "C:\VSS Carpetas de Trabajo\Mobile\"
               Else
                  strCarpetaTrabajo = "C:\VSS Carpetas de Trabajo\"
               End If
            End If
            
         Case "\\alg01\d\Software"
            If InStr(strComponente, "ALGControls") Or InStr(strComponente, "PowerMask") Then
               strCarpetaTrabajo = "\\alg01\d\Software\Controles\"
            Else
               If InStr(strComponente, "AlgInterop") Or InStr(strComponente, "AlgMobile") Then
                  strCarpetaTrabajo = "\\alg01\d\Software\Mobile\"
               Else
                  strCarpetaTrabajo = "\\alg01\d\Software\"
               End If
            End If
            
         Case "\\alg01\d\Otros\Versiones Fuentes\" & cmbVersion.List(1)
            If InStr(strComponente, "AlgInterop") Or InStr(strComponente, "AlgMobile") Then
               strCarpetaTrabajo = "\\alg01\d\Otros\Versiones Fuentes\" & cmbVersion.List(1) & "\Mobile\"
            Else
               strCarpetaTrabajo = "\\alg01\d\Otros\Versiones Fuentes\" & cmbVersion.List(1) & "\"
            End If
            
         Case Else
               If InStr(strComponente, "ALGControls") Or InStr(strComponente, "PowerMask") Then
                  strCarpetaTrabajo = cmbCarpeta.Text & "\Controles\"
               Else
                  If InStr(strComponente, "AlgInterop") Or InStr(strComponente, "AlgMobile") Then
                     strCarpetaTrabajo = cmbCarpeta.Text & "\Mobile\"
                  Else
                     strCarpetaTrabajo = cmbCarpeta.Text & "\"
                  End If
               End If
      End Select
   End If
   GetCarpetaTrabajo = strCarpetaTrabajo
End Function
Private Function W10() As Boolean
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.FolderExists(OUTDIR_X86) Then
        W10 = True
    Else
        W10 = False
    End If
End Function

Private Sub VerDllHosts()
Dim bComponenteServer As Boolean

   On Error Resume Next
   
   bComponenteServer = False
   For ix = 0 To CANTIDAD_TOTAL_DLL
      If chkComponente(ix) = vbChecked Then
         If InStr(chkComponente(ix).Caption, "DS") Or _
            InStr(chkComponente(ix).Caption, "SP") Or _
            InStr(chkComponente(ix).Caption, "DataAccess") Or _
            InStr(chkComponente(ix).Caption, "DataShare") Then
            
            bComponenteServer = True
         End If
      End If
   Next
   
   If bComponenteServer Then
      strAction = "Parando DllHosts..."
      lst1.AddItem strAction
      lst1.ListIndex = lst1.ListCount - 1
      ShutDownServerByName "", "Algoritmo"
   End If
End Sub

Private Sub ShutDownServerByName(strComputer As String, strPackage$)
Dim cat  As MTSAdmin.Catalog
Dim pkgs  As MTSAdmin.CatalogCollection
Dim pkg  As MTSAdmin.CatalogObject
Dim pkgutil  As MTSAdmin.PackageUtil
  
   On Error Resume Next
   
   Set cat = GetCatalog(strComputer)
   Set pkgs = cat.GetCollection("Packages")
   Set pkg = GetObjectFromCollection(pkgs, strPackage)
   If pkg Is Nothing Then
      Err.Raise Err.Number, Err.source, "El paquete " & strPackage & " no existe"
      Exit Sub
   End If
   
   Set pkgutil = pkgs.GetUtilInterface
   Call pkgutil.ShutdownPackage(pkg.Value("ID"))
End Sub

Private Function GetCatalog(strComputer As String) As MTSAdmin.Catalog
  Dim cat As New MTSAdmin.Catalog
  
  On Error Resume Next
  
  If strComputer <> "" Then
    cat.Connect strComputer
  End If
  Set GetCatalog = cat
End Function

Private Function GetObjectFromCollection(coll As MTSAdmin.CatalogCollection, strObjName As String) As MTSAdmin.CatalogObject
  Dim obj  As MTSAdmin.CatalogObject
  
  On Error Resume Next
  
  coll.Populate
  For Each obj In coll
    If obj.Name = strObjName Then
      Set GetObjectFromCollection = obj
      Exit Function
    End If
  Next
  Set GetObjectFromCollection = Nothing
  
End Function

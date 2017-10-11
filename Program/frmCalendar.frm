VERSION 5.00
Begin VB.Form frmCalendar 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "³¯Â¥ ¼±ÅÃ"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5295
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.TextBox txtMonth 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtYear 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   51
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   6
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "÷Ï"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   5
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "ÐÝ"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "ÙÊ"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   3
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "â©"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "ûý"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "êÅ"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "ìí"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   41
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   40
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   39
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   38
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   37
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   36
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   35
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   34
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   33
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   32
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   31
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   30
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   29
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   28
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   27
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   26
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   25
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   24
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   23
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   22
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   21
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   20
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   19
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   18
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Caption         =   "¼±ÅÃÇÏ½Ç ³¯Â¥¸¦ ´õºíÅ¬¸¯ÇÏ¼¼¿ä."
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4815
   End
   Begin VB.Image imgNext 
      Height          =   495
      Left            =   4560
      Picture         =   "frmCalendar.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   525
   End
   Begin VB.Image imgPrev 
      Height          =   495
      Left            =   240
      Picture         =   "frmCalendar.frx":13FC
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¿ù"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3960
      TabIndex        =   53
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "³â"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2400
      TabIndex        =   52
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function updateCalendar()
Dim i As Integer

For i = 0 To 41
txtDate(i) = ""
txtDate(i).Enabled = False
Next

Dim t_year As Integer
t_year = txtYear
Dim m_day As Integer
Select Case txtMonth
Case 1, 3, 5, 7, 8, 10, 12
    m_day = 31
Case 4, 6, 9, 11
    m_day = 30
Case 2
    If (t_year Mod 4 = 0 And t_year Mod 100 <> 0) Or t_year Mod 400 = 0 Then
        m_day = 29
    Else
        m_day = 28
    End If
End Select

Dim w_day As Integer
w_day = Weekday(CDate(txtYear & "-" & txtMonth & "-01")) - 2
For i = 1 To m_day
    txtDate(w_day + i).Text = i
    txtDate(w_day + i).Enabled = True
Next

End Function

Private Sub Form_Load()

Dim i As Integer
For i = 0 To 41
txtDate(i) = ""
txtDate(i).Enabled = False
Next

Dim t_year As Integer
t_year = Year(g_Date)
txtYear = t_year
txtMonth = Month(g_Date)
Dim m_day As Integer
Select Case Month(Date)
Case 1, 3, 5, 7, 8, 10, 12
    m_day = 31
Case 4, 6, 9, 11
    m_day = 30
Case 2
    If (t_year Mod 4 = 0 And t_year Mod 100 <> 0) Or t_year Mod 400 = 0 Then
        m_day = 29
    Else
        m_day = 28
    End If
End Select
Dim w_day As Integer
w_day = Weekday(CDate(Year(g_Date) & "-" & Month(g_Date) & "-" & "01")) - 2
For i = 1 To m_day
    txtDate(w_day + i).Text = i
    txtDate(w_day + i).Enabled = True
Next
End Sub

Private Sub imgNext_Click()
Dim m_Month As Integer
m_Month = txtMonth
m_Month = m_Month + 1
If m_Month > 12 Then
    Dim m_Year As Integer
    m_Month = m_Month - 12
    m_Year = txtYear
    m_Year = m_Year + 1
    txtYear = m_Year
End If

txtMonth = m_Month
Call updateCalendar
End Sub

Private Sub imgPrev_Click()
Dim m_Month As Integer
m_Month = txtMonth
m_Month = m_Month - 1
If m_Month < 1 Then
    Dim m_Year As Integer
    m_Month = m_Month + 12
    m_Year = txtYear
    m_Year = m_Year - 1
    txtYear = m_Year
End If
txtMonth = m_Month
Call updateCalendar
End Sub

Private Sub txtDate_DblClick(Index As Integer)
g_Date = CDate(txtYear & "-" & txtMonth & "-" & txtDate(Index))
Unload frmCalendar
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
txtDate(Index).BackColor = vbCyan
End Sub

Private Sub txtDate_LostFocus(Index As Integer)
txtDate(Index).BackColor = vbWhite
End Sub


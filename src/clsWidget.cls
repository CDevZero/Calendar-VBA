VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Calendar v1.0
' 
' (c) CDevZero - https://github.com/CDevZero/Calendar-VBA
'
' @class clsWidget
' @author nathandrakehs@gmail.com
' @license MIT http://www.opensource.org/licenses/mit-license.php
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

' ============================================= '
' Public Variables
' ============================================= '

Public WithEvents Widget As MSForms.Label
Attribute Widget.VB_VarHelpID = -1

' ============================================= '
' Private Variables
' ============================================= '

Private m_callback As clsEventsForWidgets
Private m_isWeekDay As Boolean
Private m_date As Date
Private m_index As Integer

' ============================================= '
' Public Methods
' ============================================= '

''
' @method Callback
' @param {clsEventsForWidgets} callbackObject
''

Public Property Set Callback(ByRef callbackObject As clsEventsForWidgets)
    Set m_callback = callbackObject
End Property

''
' @method Callback
' @return {clsEventsForWidgets} Callback
''

Public Property Get Callback() As clsEventsForWidgets
    Set Callback = m_callback
End Property

''
' @method Value
' @param {Date} DateValue
''

Public Property Let Value(ByVal DateValue As Date)
    m_date = DateValue
End Property

''
' @method Value
' @return {Date} Value
''

Public Property Get Value() As Date
    Value = m_date
End Property

''
' @method Index
' @param {Integer} IndexValue
''

Public Property Let Index(ByVal IndexValue As Integer)
    m_index = IndexValue
End Property

''
' @method Index
' @return {Integer} Index
''

Public Property Get Index() As Integer
    Index = m_index
End Property

''
' @method IsWeekDay
' @param {Boolean} WeekDayStatus
''

Public Property Let IsWeekDay(ByVal WeekDayStatus As Boolean)
    m_isWeekDay = WeekDayStatus
End Property

' ============================================= '
' Private Events
' ============================================= '

Private Sub Widget_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If m_isWeekDay Then Exit Sub

    If (Not m_callback Is Nothing) Then
        m_callback.DblClick m_date
    End If
End Sub

Private Sub Widget_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If m_isWeekDay Then Exit Sub
    
    If (Not m_callback Is Nothing) Then
        m_callback.MouseMove m_index, X, Y
    End If
End Sub

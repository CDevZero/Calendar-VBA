VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Calendar v1.0
' 
' (c) CDevZero - https://github.com/CDevZero/Calendar-VBA
'
' `Calendar` returns the selected date on the created form.
'
' @usage example:
' ```VB.net
' Dim Calendar As New clsCalendar
' Dim SelectedDate As Date
' 
' SelectedDate = Calendar.getDate '( returns date in format `dd.mm.yyyy` )
' ' or SelectDate = Calendar
'
' Set Calendar = Nothing
' ```
'
' @class clsCalendar
' @author nathandrakehs@gmail.com
' @license MIT http://www.opensource.org/licenses/mit-license.php
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

' ============================================= '
' Constants, Enumerations and Private Variables
' ============================================= '

Implements clsEventsForWidgets

Private m_widgets() As New clsWidget

Private m_globalDate As Date
Private m_returnDate As Date
Private m_startPoint As Date

Private m_previousIndex As Integer
Private m_startPointIndex As Integer

Private m_mode As Long
Private m_previousMode As Long

Private m_component As Object
Private m_calendar As Object

Private m_caption As String
Private m_firstDayOfTheWeek As Long

Private WithEvents m_leftSlider As MSForms.Label
Attribute m_leftSlider.VB_VarHelpID = -1
Private WithEvents m_rightSlider As MSForms.Label
Attribute m_rightSlider.VB_VarHelpID = -1
Private WithEvents m_switchMode As MSForms.Label
Attribute m_switchMode.VB_VarHelpID = -1

Private Const FORM_CAPTION As String = "Calendar"
Private Const FORM_HEIGHT As Integer = 150
Private Const FORM_WIDTH As Integer = 150

Private Const START_POS_X As Integer = 12
Private Const START_POS_Y As Integer = 30

Private Enum InterfaceColor
    Default = &H8000000F
    Selected = &HFFFF80
    CurrentDate = &HFFFFC0
    Inactive = &H8000000C
    FontDefault = &H80000012
    FontSelected = &H8000000D
End Enum

Public Enum vbCalendarMode
    vbModeIsDay = 0
    vbModeIsMonth = 1
    vbModeIsYear = 2
End Enum

Private Enum vbExcludeController
    vbAllControllers = -1
    vbLeftSlider = 0
    vbSwitchMode = 1
    vbRightSlider = 2
End Enum

' ============================================= '
' Initialize and Terminate class
' ============================================= '

Private Sub Class_Initialize()
    m_caption = FORM_CAPTION
    m_firstDayOfTheWeek = vbUseSystemDayOfWeek
    Call ResetIndex
End Sub

Private Sub Class_Terminate()
    Erase m_widgets

    ThisWorkbook.VBProject.VBComponents.Remove m_component

    Set m_leftSlider = Nothing
    Set m_rightSlider = Nothing
    Set m_switchMode = Nothing
    Set m_calendar = Nothing
    Set m_component = Nothing
End Sub

' ============================================= '
' Public Methods
' ============================================= '

''
' The method returns the selected date.
'
' @usage example:
' ```VB.Net
' Dim Calendar As New clsCalendar
' Dim SelectedDate As Date

' SelectedDate = Calendar.getDate(vbModeIsMonth, #05/01/2020#)
' ' or SelectDate = Calendar(StartPoint:=#05/01/2020#)
' 
' Set Calendar = Nothing
' ```
'
' @method getDate
' @param {Long} Mode (optional)
' @param {Date} StartPoint (optional)
' @return {Date} dd.mm.yyyy
''

Public Function getDate(Optional Mode As vbCalendarMode = vbModeIsDay, Optional StartPoint As Date) As Date
Attribute getDate.VB_UserMemId = 0
    m_mode = Mode

    If m_mode = vbModeIsMonth Then
        m_previousMode = vbModeIsYear
    End If

    If StartPoint = CDate(0) Then
        m_startPoint = Now
    Else
        m_startPoint = StartPoint
    End If

    m_globalDate = Format(m_startPoint, "dd.mm.yyyy")
    
    Call CreateCalendar
    Call ShowCalendar

    getDate = m_returnDate
    
    Call Class_Terminate
End Function

''
' The method sets the caption of the form.
'
' @usage example:
' ```VB.net
' Dim Calendar As New clsCalendar
'
' Calendar.Caption = "Calendar"
' ```
'
' @method Caption
' @param {String} StrValue
''

Public Property Let Caption(StrValue As String)
    m_caption = StrValue
End Property

''
' The method sets the first day of the week.
'
' @usage example:
' ```VB.net
' Dim Calendar As New clsCalendar
'
' Calendar.FirstDayOfTheWeek = vbSunday
' ```
'
' @method FirstDayOfTheWeek
' @param {Long} FirstDay
''

Public Property Let FirstDayOfTheWeek(FirstDay as vbDayOfWeek)
    m_firstDayOfTheWeek = FirstDay
End Property

' ============================================= '
' Private Methods
' ============================================= '

''
' The method reset indexes.
'
' @method DropIndex
''

Private Sub ResetIndex()
    m_previousIndex = -1
    m_startPointIndex = -1
End Sub

''
' The method creates a custom form in VBA Project.
'
' @method CreateCalendar
' @throws
''

Private Sub CreateCalendar()
    Set m_component = ThisWorkbook.VBProject.VBComponents.Add(3)

    With m_component
        .Properties("Caption") = m_caption
        .Properties("Height") = FORM_HEIGHT
        .Properties("Width") = FORM_WIDTH
    End With

    Set m_calendar = VBA.UserForms.Add(m_component.Name)

    With m_calendar.Controls
        Set m_leftSlider = .Add("Forms.Label.1", "LeftSlider", True)
        Set m_rightSlider = .Add("Forms.Label.1", "RightSlider", True)
        Set m_switchMode = .Add("Forms.Label.1", "SwitchButton", True)
    End With

    With m_leftSlider
        .Caption = "<"
        .TextAlign = fmTextAlignCenter
        .Width = 16
        .Height = 12
        .Left = START_POS_X
        .Top = START_POS_Y - 20
    End With

    With m_rightSlider
        .Caption = ">"
        .TextAlign = fmTextAlignRight
        .Width = 16
        .Height = 12
        .Left = START_POS_X + .Width * 6
        .Top = START_POS_Y - 20
    End With

    With m_switchMode
        .TextAlign = fmTextAlignCenter
        .Width = 60
        .Height = 12
        .Left = START_POS_X + (m_rightSlider.Left + m_rightSlider.Width * 1.5 - m_leftSlider.Left) / 2 - .Width / 2 - 2
        .Top = START_POS_Y - 20
    End With
End Sub

''
' The method displays a custom form.
'
' @method ShowCalendar
' @param {integer} Modal (optional)
' @throws
''

Private Sub ShowCalendar(Optional ByVal Modal As Integer = 1)
    Select Case m_mode
        Case vbModeIsDay
            Call CreateDaysWidgets
        Case vbModeIsMonth
            Call CreateMonthsWidgets
        Case vbModeIsYear
            Call CreateYearsWidgets
    End Select

    m_calendar.Show
End Sub

''
' The method deletes a custom form..
'
' @method DeleteCalendar
''

Private Sub DeleteCalendar()
    Unload m_calendar
    Call Class_Terminate
End Sub

''
' The method creates Days Widgets in a custom form.
'
' @method CreateDaysWidgets
''

Private Sub CreateDaysWidgets()
    Dim startDate As Date
    Dim i As Integer

    ReDim m_widgets(0 To 48)

    startDate = Format(m_globalDate - Format(m_globalDate, "d") + 1, "dd.mm.yyyy")
    startDate = Format(startDate - Format(startDate, "w", m_firstDayOfTheWeek) + 1, "dd.mm.yyyy")

    m_switchMode.Caption = MonthName(Format(m_globalDate, "m")) & " " & Format(m_globalDate, "yyyy")

    For i = LBound(m_widgets) To UBound(m_widgets)
        Set m_widgets(i).Widget = m_calendar.Controls.Add("Forms.Label.1", "Label" & i + 4, True)
        Set m_widgets(i).Callback = Me

        m_widgets(i).Index = i

        With m_widgets(i).Widget
            .Width = 16
            .Height = 12
            .BorderStyle = fmBorderStyleNone

            Select Case i
                Case 0 To 6
                    .Caption = WeekdayName(i + 1, True, m_firstDayOfTheWeek)
                    .TextAlign = fmTextAlignRight
                    .Left = START_POS_X + .Width * i
                    .Top = START_POS_Y

                    m_widgets(i).IsWeekDay = True
                Case 7 To 48
                    m_widgets(i).Value = Format(startDate + i - 7, "dd.mm.yyyy")

                    If m_widgets(i).Value = Format(m_startPoint, "dd.mm.yyyy") Then
                        .BackColor = InterfaceColor.CurrentDate
                    End If

                    If Format(m_widgets(i).Value, "mm") <> Format(m_globalDate, "mm") Then
                        .ForeColor = InterfaceColor.Inactive
                    End If

                    .Left = START_POS_X + .Width * (i Mod 7)
                    .Top = START_POS_Y + .Height * (i \ 7)
                    .Caption = Format(m_widgets(i).Value, "d")
                    .TextAlign = fmTextAlignRight
            End Select
        End With
    Next i

    m_calendar.Controls.Add "Forms.Label.1", "Label53", True

    With m_calendar.Controls("Label53")
        .Top = START_POS_Y + 11
        .Left = START_POS_X
        .Width = 117
        .Height = 1
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = InterfaceColor.Inactive
    End With
End Sub

''
' The method creates Months Widgets in a custom form.
'
' @method CreateMonthsWidgets
''

Private Sub CreateMonthsWidgets()
    Dim i As Integer

    m_switchMode.Caption = Format(m_globalDate, "yyyy")

    ReDim m_widgets(0 To 11)

    For i = LBound(m_widgets) To UBound(m_widgets)
        Set m_widgets(i).Widget = m_calendar.Controls.Add("Forms.Label.1", "Label" & i + 4, True)
        Set m_widgets(i).Callback = Me

        m_widgets(i).Value = Format(i + 1 & "." & m_switchMode.Caption, "mm.yyyy")
        m_widgets(i).Index = i

        With m_widgets(i).Widget
            .Caption = vbLf & MonthName(Format(m_widgets(i).Value, "m"), True)

            If m_widgets(i).Value = Format(m_startPoint, "mm.yyyy") Then
                .BackColor = InterfaceColor.CurrentDate
            End If

            If m_widgets(i).Value = Format(m_globalDate, "mm.yyyy") Then
                .BorderColor = InterfaceColor.Selected
                .BorderStyle = fmBorderStyleSingle
            End If

            .Width = 28
            .Height = 28
            .Left = START_POS_X + .Width * (i Mod 4)
            .Top = START_POS_Y + .Height * (i \ 4)
            .TextAlign = fmTextAlignCenter
        End With
    Next i
End Sub

''
' The method creates Years Widgets in a custom form.
'
' @method CreateYearsWidgets
''

Private Sub CreateYearsWidgets()
    Dim i As Integer, d As Integer, m As Integer, year As Integer

    d = Format(m_globalDate, "dd")
    m = Format(m_globalDate, "mm")

    If d = 29 And m = 2 Then d = 28

    year = Format(m_globalDate, "yyyy")

    m_switchMode.Caption = year - (year Mod 9) & "-" & year + (8 - (year Mod 9))

    ReDim m_widgets(0 To 8)

    For i = LBound(m_widgets) To UBound(m_widgets)
        Set m_widgets(i).Widget = m_calendar.Controls.Add("Forms.Label.1", "Label" & i + 4, True)
        Set m_widgets(i).Callback = Me

        m_widgets(i).Value = CDate(d & "." & m & "." & (year - (year Mod 9)) + i)
        m_widgets(i).Index = i

        With m_widgets(i).Widget
            .Caption = vbLf & Format(m_widgets(i).Value, "yyyy")

            If Format(m_widgets(i).Value, "yyyy") = Format(m_startPoint, "yyyy") Then
                .BackColor = InterfaceColor.CurrentDate
            End If

            If Format(m_widgets(i).Value, "yyyy") = Format(m_globalDate, "yyyy") Then
                .BorderColor = InterfaceColor.Selected
                .BorderStyle = fmBorderStyleSingle
            End If

            .Width = 38
            .Height = 28
            .Left = START_POS_X + .Width * (i Mod 3)
            .Top = START_POS_Y + .Height * (i \ 3)
            .TextAlign = fmTextAlignCenter
        End With
    Next i
End Sub

''
' The method sets the default color for the previous widget.
'
' @method SetColorForPreviousWidget
' @param {Integer} WidgetIndex
' @param {Single} X
' @param {Single} Y
''

Private Sub SetColorForPreviousWidget(ByVal WidgetIndex As Integer, ByVal X As Single, ByVal Y As Single)
    Const STEP As Integer = 3

    Dim xMax As Integer
    Dim yMax As Integer

    With m_widgets(WidgetIndex).Widget
        xMax = .Width
        yMax = .Height
    End With

    If ((X >= 0 And X <= STEP) And (Y >= 0 And Y <= yMax)) Then
        GoTo SetColor
    ElseIf ((X >= xMax - STEP And X <= xMax) And (Y >= 0 And Y <= yMax)) Then
        GoTo SetColor
    ElseIf ((X >= 0 And X <= xMax) And (Y >= 0 And Y <= STEP)) Then
        GoTo SetColor
    ElseIf ((X >= 0 And X <= xMax) And (Y >= yMax - STEP And Y <= yMax)) Then
        GoTo SetColor
    End If

    Exit Sub

SetColor:
    If m_previousIndex = m_startPointIndex Then
        m_widgets(m_previousIndex).Widget.BackColor = InterfaceColor.CurrentDate
    Else
        m_widgets(m_previousIndex).Widget.BackColor = InterfaceColor.Default
    End If
End Sub

''
' The method sets the default color of the control.
'
' @method SetDefaultColorForControllers
' @param {Long} Exclude
''

Private Sub SetDefaultColorForControllers(Optional Exclude As vbExcludeController = vbAllControllers)
    Select Case Exclude
        Case vbAllControllers
            m_leftSlider.ForeColor = InterfaceColor.FontDefault
            m_switchMode.ForeColor = InterfaceColor.FontDefault
            m_rightSlider.ForeColor = InterfaceColor.FontDefault
        Case vbLeftSlider
            m_switchMode.ForeColor = InterfaceColor.FontDefault
            m_rightSlider.ForeColor = InterfaceColor.FontDefault
        Case vbSwitchMode
            m_leftSlider.ForeColor = InterfaceColor.FontDefault
            m_rightSlider.ForeColor = InterfaceColor.FontDefault
        Case vbRightSlider
            m_leftSlider.ForeColor = InterfaceColor.FontDefault
            m_switchMode.ForeColor = InterfaceColor.FontDefault
    End Select
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Function DateOffset(ByVal DateValue As Date, ByVal Offset as integer) As Date
    DateOffset = Format(Format(DateValue, "d") & "." & Format(DateValue, "mm") & "." & Format(DateValue, "yyyy") + offset, "dd.mm.yyyy")
End Function

' ============================================= '
' Private Events
' ============================================= '

Private Sub clsEventsForWidgets_DblClick(ByVal ReturnDate As Date)
    Select Case m_mode
        Case vbModeIsDay
            m_returnDate = ReturnDate
            Call DeleteCalendar
        Case vbModeIsMonth
            m_globalDate = ReturnDate
            m_previousMode = 2
            Call m_switchMode_Click
        Case vbModeIsYear
            m_globalDate = ReturnDate
            Call m_switchMode_Click
    End Select
End Sub

Private Sub clsEventsForWidgets_MouseMove(ByVal WidgetIndex As Integer, ByVal X As Single, ByVal Y As Single)
    Call SetDefaultColorForControllers

    If m_previousIndex = -1 Then m_previousIndex = WidgetIndex

    If m_widgets(WidgetIndex).Widget.BackColor = InterfaceColor.CurrentDate Then
        m_widgets(m_previousIndex).Widget.BackColor = InterfaceColor.Default
        m_widgets(WidgetIndex).Widget.BackColor = InterfaceColor.Selected
        m_startPointIndex = WidgetIndex
        Exit Sub
    End If

    m_widgets(WidgetIndex).Widget.BackColor = InterfaceColor.Selected

    If WidgetIndex <> m_previousIndex Then
        If m_previousIndex = m_startPointIndex Then
            m_widgets(m_previousIndex).Widget.BackColor = InterfaceColor.CurrentDate
        Else
            m_widgets(m_previousIndex).Widget.BackColor = InterfaceColor.Default
        End If

        m_previousIndex = WidgetIndex
    End If

    Select Case m_mode
        Case vbModeIsDay
            Select Case WidgetIndex
                Case 7 To 14, 20 To 21, 27 To 28, 34 To 35, 41 To 48
                    Call SetColorForPreviousWidget(WidgetIndex, X, Y)
            End Select
        Case vbModeIsMonth
            Select Case WidgetIndex
                Case 0 To 4, 7 To 11
                    Call SetColorForPreviousWidget(WidgetIndex, X, Y)
            End Select
        Case vbModeIsYear
            Call SetColorForPreviousWidget(WidgetIndex, X, Y)
    End Select
End Sub

Private Sub m_switchMode_Click()
    Dim i As Integer

    Select Case m_mode
        Case vbModeIsDay
            For i = LBound(m_widgets) To UBound(m_widgets) + 1
                m_calendar.Controls.Remove "Label" & i + 4
            Next i

            Erase m_widgets
            Call CreateMonthsWidgets

            m_previousMode = vbModeIsDay
            m_mode = vbModeIsMonth
        Case vbModeIsMonth
            For i = LBound(m_widgets) To UBound(m_widgets)
                m_calendar.Controls.Remove "Label" & i + 4
            Next i

            Erase m_widgets

            Select Case m_previousMode
                Case vbModeIsDay
                    Call CreateYearsWidgets
                    m_mode = vbModeIsYear
                Case vbModeIsYear
                    Call CreateDaysWidgets
                    m_mode = vbModeIsDay
            End Select
        Case vbModeIsYear
            For i = LBound(m_widgets) To UBound(m_widgets)
                m_calendar.Controls.Remove "Label" & i + 4
            Next i
            
            Erase m_widgets
            Call CreateMonthsWidgets
            
            m_previousMode = vbModeIsYear
            m_mode = vbModeIsMonth
    End Select
    
    Call Class_Initialize
End Sub

Private Sub m_leftSlider_Click()
    Dim i As Integer

    Select Case m_mode
        Case vbModeIsDay
            Do While MONTH(m_globalDate - i) = MONTH(m_globalDate)
                i = i + 1
            Loop

            m_globalDate = m_globalDate - i
            
            For i = LBound(m_widgets) To UBound(m_widgets) + 1
                m_calendar.Controls.Remove "Label" & i + 4
            Next i

            Erase m_widgets
            Call CreateDaysWidgets
        Case vbModeIsMonth
            m_globalDate = DateOffset(m_globalDate, -1)

            For i = LBound(m_widgets) To UBound(m_widgets)
                m_calendar.Controls.Remove "Label" & i + 4
            Next i

            Erase m_widgets
            Call CreateMonthsWidgets
        Case vbModeIsYear
            m_globalDate = DateOffset(m_globalDate, -9)

            For i = LBound(m_widgets) To UBound(m_widgets)
                m_calendar.Controls.Remove "Label" & i + 4
            Next i

            Erase m_widgets
            Call CreateYearsWidgets
    End Select

    Call Class_Initialize
End Sub

Private Sub m_rightSlider_Click()
    Dim i As Integer

    Select Case m_mode
        Case vbModeIsDay
            Do While MONTH(m_globalDate + i) = MONTH(m_globalDate)
                i = i + 1
            Loop

            m_globalDate = m_globalDate + i

            For i = LBound(m_widgets) To UBound(m_widgets) + 1
                m_calendar.Controls.Remove "Label" & i + 4
            Next i

            Erase m_widgets
            Call CreateDaysWidgets
        Case vbModeIsMonth
            m_globalDate = DateOffset(m_globalDate, 1)

            For i = LBound(m_widgets) To UBound(m_widgets)
                m_calendar.Controls.Remove "Label" & i + 4
            Next i

            Erase m_widgets
            Call CreateMonthsWidgets
        Case vbModeIsYear
            m_globalDate = DateOffset(m_globalDate, 9)

            For i = LBound(m_widgets) To UBound(m_widgets)
                m_calendar.Controls.Remove "Label" & i + 4
            Next i

            Erase m_widgets
            Call CreateYearsWidgets
    End Select

    Call Class_Initialize
End Sub

Private Sub m_leftSlider_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call SetDefaultColorForControllers(vbLeftSlider)
    Call Controller_MouseMove(m_leftSlider, X, Y)
End Sub

Private Sub m_switchMode_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call SetDefaultColorForControllers(vbSwitchMode)
    Call Controller_MouseMove(m_switchMode, X, Y)
End Sub

Private Sub m_rightSlider_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call SetDefaultColorForControllers(vbRightSlider)
    Call Controller_MouseMove(m_rightSlider, X, Y)
End Sub

Private Sub Controller_MouseMove(ByRef objControl As MSForms.Label, ByVal X As Single, ByVal Y As Single)
    Const STEP As Integer = 3

    Dim xMax As Integer
    Dim yMax As Integer

    xMax = objControl.Width
    yMax = objControl.Height
    objControl.ForeColor = InterfaceColor.FontSelected

    If ((X >= 0 And X <= STEP) And (Y >= 0 And Y <= yMax)) Then
        GoTo SetColor
    ElseIf ((X >= xMax - STEP And X <= xMax) And (Y >= 0 And Y <= yMax)) Then
        GoTo SetColor
    ElseIf ((X >= 0 And X <= xMax) And (Y >= 0 And Y <= STEP)) Then
        GoTo SetColor
    ElseIf ((X >= 0 And X <= xMax) And (Y >= yMax - STEP And Y <= yMax)) Then
        GoTo SetColor
    End If

    Exit Sub

SetColor:
    objControl.ForeColor = InterfaceColor.FontDefault
End Sub
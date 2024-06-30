Attribute VB_Name = "ModCharts"
Option Explicit

Sub addcharts()
Charts.Add
Sheets.Add Type:=XlSheetType.xlChart
End Sub
Sub deletecharts()
Application.DisplayAlerts = False
Charts.Delete
Application.DisplayAlerts = True
End Sub



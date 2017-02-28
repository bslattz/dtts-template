Attribute VB_Name = "Module1"
Option Explicit
Public vstoQuery As WeekEndingTabs.QueryWeekending

Sub printState(status As String, guard As Boolean)
  Set vstoQuery = GetManagedClass(ThisWorkbook)
  Debug.Print status & IIf(guard, "guarded", "NOT guarded")
  Debug.Print vstoQuery.DisplayTaggedSheets
  Debug.Print Format(vstoQuery.DisplayDates, "dd/mm/yyy")
End Sub
Sub Log(message As String)
  Debug.Print message
End Sub


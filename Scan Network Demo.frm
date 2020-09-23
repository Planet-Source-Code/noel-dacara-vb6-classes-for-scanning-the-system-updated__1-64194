VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents NET As cScanNetwork
Attribute NET.VB_VarHelpID = -1

Private Sub Form_Click()
    NET.BeginScanning
End Sub

Private Sub Form_Load()
    Set NET = New cScanNetwork
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set NET = Nothing
End Sub

Private Sub NET_CurrentComputer(Name As String, Domain As String)
    Print "Computer Scanned:"
    Print vbTab; "Name = "; Name
    Print vbTab; "Domain = "; Domain
End Sub

Private Sub NET_CurrentDomain(Name As String, Provider As String)
    Print "Domain Scanned:"
    Print vbTab; "Name = "; Name
    Print vbTab; "Provider = "; Provider
End Sub

Private Sub NET_DoneScanning(TotalDomains As Long, TotalComputers As Long)
    Print "Done Scanning:"
    Print vbTab; "Total Domains = "; TotalDomains
    Print vbTab; "Total Computers = "; TotalComputers
End Sub

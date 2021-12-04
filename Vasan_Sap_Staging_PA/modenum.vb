Imports System.Data.SqlClient
Module modEnum
    Public definenew As Boolean = False
    Public MServerName As String
    Public MDBName As String
    Public MUID As String
    Public MPWD As String
    Public MLoginType As String
    Public MSAPUID As String
    Public MSAPPWd As String
    Public cmd As SqlCommand
    Public con_staging As New SqlConnection
    Public con_SAP As New SqlConnection
    Public da As SqlDataAdapter
    Public dt As DataTable
    Public strsql As String
    Public command_sap As SqlCommand
    Public command_staging As SqlCommand
End Module


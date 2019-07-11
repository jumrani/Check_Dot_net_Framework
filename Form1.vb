Imports System
Imports Microsoft.Win32
Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim V1 As String = checkdotnetversion()
    End Sub
    Public Function checkdotnetversion() As String
        Dim maxDotNetVersion As String = GetVersionFromRegistry()
        If String.Compare(maxDotNetVersion, "4.5") >= 0 Then
            Dim v45Plus As String = Get45PlusFromRegistry()
            If v45Plus <> "" Then maxDotNetVersion = v45Plus
        End If
        If maxDotNetVersion <> "" Then
            maxDotNetVersion = "*** Maximum .NET version number found is: " & maxDotNetVersion & "***"
        Else
            maxDotNetVersion = ".NET FRAMEWORK IS NOT FOUND"
        End If
        Return maxDotNetVersion
    End Function
    Public Function Get45PlusFromRegistry() As String
        Dim dotNetVersion As String = ""
        Const subkey As String = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"

        Using ndpKey As RegistryKey = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, "").OpenSubKey(subkey)

            If ndpKey IsNot Nothing AndAlso ndpKey.GetValue("Release") IsNot Nothing Then
                dotNetVersion = CheckFor45PlusVersion(CInt(ndpKey.GetValue("Release")))
            Else
                dotNetVersion = ".NET Framework Version 4.5 or later is not detected."
            End If
        End Using
        Return dotNetVersion
    End Function
    Public Function CheckFor45PlusVersion(ByVal releaseKey As Integer) As String
        If releaseKey >= 528040 Then Return "4.8 or later"
        If releaseKey >= 461808 Then Return "4.7.2"
        If releaseKey >= 461308 Then Return "4.7.1"
        If releaseKey >= 460798 Then Return "4.7"
        If releaseKey >= 394802 Then Return "4.6.2"
        If releaseKey >= 394254 Then Return "4.6.1"
        If releaseKey >= 393295 Then Return "4.6"
        If (releaseKey >= 379893) Then Return "4.5.2"
        If (releaseKey >= 378675) Then Return "4.5.1"
        If (releaseKey >= 378389) Then Return "4.5"
        Return "No 4.5 or later version detected"
    End Function
    Public Function GetVersionFromRegistry() As String
        Dim maxDotNetVersion As String = ""

        Using ndpKey As RegistryKey = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, "").OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\")

            For Each versionKeyName As String In ndpKey.GetSubKeyNames()

                If versionKeyName.StartsWith("v") Then
                    Dim versionKey As RegistryKey = ndpKey.OpenSubKey(versionKeyName)
                    Dim name As String = CStr(versionKey.GetValue("Version", ""))
                    Dim sp As String = versionKey.GetValue("SP", "").ToString()
                    Dim install As String = versionKey.GetValue("Install", "").ToString()

                    If install = "" Then
                        If String.Compare(maxDotNetVersion, name) < 0 Then maxDotNetVersion = name
                    Else

                        If sp <> "" AndAlso install = "1" Then
                            If String.Compare(maxDotNetVersion, name) < 0 Then maxDotNetVersion = name
                        End If
                    End If

                    If name <> "" Then
                        Continue For
                    End If

                    For Each subKeyName As String In versionKey.GetSubKeyNames()
                        Dim subKey As RegistryKey = versionKey.OpenSubKey(subKeyName)
                        name = CStr(subKey.GetValue("Version", ""))

                        If name <> "" Then
                            sp = subKey.GetValue("SP", "").ToString()
                        End If

                        install = subKey.GetValue("Install", "").ToString()

                        If install = "" Then
                            If String.Compare(maxDotNetVersion, name) < 0 Then maxDotNetVersion = name
                        Else

                            If sp <> "" AndAlso install = "1" Then
                                If String.Compare(maxDotNetVersion, name) < 0 Then maxDotNetVersion = name
                            ElseIf install = "1" Then
                                If String.Compare(maxDotNetVersion, name) < 0 Then maxDotNetVersion = name
                            End If
                        End If
                    Next
                End If
            Next
        End Using
        Return maxDotNetVersion
    End Function
End Class


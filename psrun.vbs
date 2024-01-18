'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                                                                 '
'     FILENAME.....: psrun.vbs                                                                    '
'     REPOSITORY...: https://github.com/lazux/psrun.git                                           '
'     AUTHOR.......: Kerim Mutlu (lazux>at>gmx>dot>net)                                           '
'     COMPANY......: Private                                                                      '
'     DATE.........: 2019-09-15                                                                   '
'     VERSION......: 0.18                                                                         '
'     SYNOPSIS.....: Runs a PowerShell-Script while suppressing the implied console window        '
'     DESCRIPTION..: This script is designed to meet the following requirements                   '
'                        * Usability with dedicated board resources or widely used tools          '
'                        * Capability to hide/show the Powershell console and run target          '
'                          Powershell scripts with elevated privileges                            '
'                        * Acceptance of a variable number of parameters, forwarding them         '
'                          to the target Powershell script according to type                      '
'     PARAMETER....: /Elevate   Executes with elevated permissions                                '
'                       /Hide   Suppresses the PowerShell console window                          '
'                       /File   Path to the PowerShell script (relative or absolute)              '
'                        * All other parameters are forwarded to the called Powershell script     '
'                          in a type-appropriate format                                           '
'     USAGE........: wscript psrun.vbs /File:{PATH} ...                                           '
'     EXAMPLES.....: wscript psrun.vbs /Command:"calc.exe" /Hide                                  '
'                    wscript psrun.vbs /NoExit /Command:"Get-ChildItem C:\Windows"                '
'                    wscript psrun.vbs /File:"C:\foo\bar.ps1" /Elevate                            '
'                    wscript psrun.vbs /File:"bar.ps1" /Voo:2.5 /Voo:"baz" /Voo:7 /Doo            '
'     LICENSE......: GNU General Public License, Version 3.0                                      '
'                                                                                                 '
'     Copyright (c) 2019 Kerim Mutlu                                                              '
'                                                                                                 '
'     This program is free software: you can redistribute and/or modify it under the terms of     '
'     the GNU General Public License as published by the Free Software Foundation, either ver     '
'     sion 3 of the License, or at your option any later version. Its distributed in the hope     '
'     that it will be useful, but WITHOUT ANY WARRANTY;  without even the implied warranty of     '
'     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License     '
'     for more details.                                                                           '
'                                                                                                 '
'     You should have received a copy of the GNU                                                  '
'     General Public License along with this program.                                             '
'     If not, see <http://www.gnu.org/licenses/>.                                                 '
'                                                                                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit



Function Main()
    Dim Space : Space = Chr(32)
    Dim View : View = 5
    Dim Verb : Verb = Null
    Dim Command : Command = "powershell.exe"
    Dim Options : Options = "-NoLogo -NonInteractive -ExecutionPolicy ByPass"
    Dim Arguments : Set Arguments = CreateObject("Scripting.Dictionary")
    Dim Result : Result = Null

    If Collect(Arguments) Then
        If Arguments.Exists("File") Then
            Dim File
            File = Resolve(Arguments.Item("File"))
            If IsNullorEmpty(File) Then
                Result = "File not found: " & Arguments.Item("File")
            Else
                Arguments.Item("File") = File
            End If
        End If

        If IsNullorEmpty(Result) Then
            Dim Line

            If Arguments.Exists("Elevate") Then
                Verb = "runas"
                Arguments.Remove("Elevate")
            End If

            If Arguments.Exists("Hide") Then
                View = 0
                Arguments.Remove("Hide")
            End If

            Line = Strcmd(Arguments)

            If Not IsNullorEmpty(Line) Then
                If Not IsNullorEmpty(Options) Then
                    Line = Options & Space & Line
                End If
                
                Dim Sha
                Set Sha = CreateObject("Shell.Application")
                Sha.ShellExecute Command, Line, Null, Verb, View
                Set Sha = Nothing
                Arguments.RemoveAll()
            Else
                Result = "Missing parameter: Specify a powershell file or command"
            End If
        End If
    Else
        Result = "No parameters were specified"
    End If

    Set Arguments = Nothing
    Main = Result
End Function



Function Collect(ByRef Arguments)
    Dim Space : Space = Chr(32)
    Dim Specifier : Specifier = "/"
    Dim Delimiter : Delimiter = ":"
    Dim Name
    Dim Value
    Dim Position
    Dim Arg
    Dim Index
    Dim Result

    Arguments.CompareMode = vbTextCompare
    
    If WScript.Arguments.Count > 0 Then
        For Each Arg In WScript.Arguments
            Name = Null
            Value = Null
            If InStr(1, Arg, Specifier) = 1 Then
                Position = InStr(3, Arg, Delimiter)
                If Position > 0 Then
                    Name = Mid(Arg, 2, Position-2)
                    Value = Convert(Mid(Arg, Position+1))

                    If Not Arguments.Exists(Name) Then
                        Arguments.Add Name, Value
                    Else
                        Dim Temp
                        If IsArray(Arguments.Item(Name)) Then
                            Temp = Arguments.Item(Name)
                            ReDim Preserve Temp(UBound(Temp)+1)
                            Temp(UBound(Temp)) = Value
                        Else
                            Temp = Array()
                            ReDim Temp(1)
                            Temp(0) = Arguments.Item(Name)
                            Temp(1) = Value
                        End If
                        Arguments(Name) = Temp
                        Erase Temp
                    End If

                Else
                    Name = Mid(Arg, 2)
                    If Not Arguments.Exists(Name) Then
                        Arguments.Add Name, Null
                    End If
                End if
            Else
                Value = Convert(Arg)
                Index = 0
                Do
                    Index = Index + 1
                Loop While Arguments.Exists(Index)
                Arguments.Add Index, Value
            End If
        Next
        Result = True
    Else
        Result = False
    End If

    Collect = Result
End Function



Function Strcmd(ByRef Arguments)
    Dim Space : Space = Chr(32)
    Dim Index : Index = Arguments.Count
    Dim Arg
    Dim Result

    For Each Arg In Arguments
        Index = Index - 1
        Result = Result & Outline(Arg, Arguments.Item(Arg))
        If Index > 0 Then
            Result = Result & Space
        End If
    Next

    Strcmd = Result
End Function



Function IsNullorEmpty(Value)
    Dim Result

    If Len("" & Value) = 0 Then
        Result = True
    Else
        Result = False
    End If

    IsNullorEmpty = Result
End Function



Function Resolve(Path)
    Dim File
    Dim Location
    Dim Result
    Dim Fso : Set Fso = CreateObject("Scripting.FileSystemObject")
    
    If Fso.FileExists(Path) Then
        Set Location = Fso.GetFile(Path)
        File = Location.Path
        Set Location = Nothing
        Result = File
    Else
        Location = Fso.GetParentFolderName(WScript.ScriptFullName)
        File = Fso.BuildPath(Location, Path)
        If Fso.FileExists(File) Then
            Result = Fso.GetAbsolutePathName(File)
        Else
            Dim Wsh
            Set Wsh = CreateObject("WScript.Shell")
            Location = Wsh.Environment("Process").Item("UserProfile")
            File = Fso.BuildPath(Location, Path)
            Set Wsh = Nothing
            If Fso.FileExists(File) Then
                Result = Fso.GetAbsolutePathName(File)
            Else
                Result = Null
            End If
        End If
    End If

    Set Fso = Nothing
    Resolve = Result
End Function



Function Convert(Value)
    Dim Result

    If IsNumeric(Value) Then
        If (CLng(Value) & "") = (Value & "") Then
            Result = CLng(Value)
        Else
            Result = CDbl(Value)
        End If
    Else
        Result = Value
    End If

    Convert = Result
End Function



Function Outline(Key, Value)
    Dim Space : Space = Chr(32)
    Dim Quote : Quote = Chr(34)
    Dim Specifier : Specifier = "-"
    Dim Delimiter : Delimiter = ","
    Dim Result

    If Not IsNumeric(Key) Then
        If Not IsNull(Value) Then
            If IsNumeric(Value) Then
                Result = Specifier & Key & Space & Value
            Else
                If IsArray(Value) Then
                    Result = Specifier & Key & Space
                    Dim Size : Size = UBound(Value)
                    Dim Index
                    For Index = 0 To Size
                        If IsNumeric(Value(Index)) Then
                            Result = Result & Value(Index)
                        Else
                            Result = Result & Quote & Value(Index) & Quote
                        End If
                        If Index < Size Then
                            Result = Result & Delimiter
                        End If
                    Next
                Else
                    Result = Specifier & Key & Space & Quote & Value & Quote
                End If
            End If
        Else
            Result = Specifier & Key
        End If
    Else
        If IsNumeric(Value) Then
            Result = Value
        Else
            Result = Quote & Value & Quote
        End If
    End If

    Outline = Result
End Function





Dim Result
Dim Locale

Locale = GetLocale()
SetLocale(2057)

Result = Main()

If IsNullorEmpty(Result) Then
    Result = 0
Else
    WScript.Echo Result
    Result = 1
End If

SetLocale(Locale)

WScript.Quit(Result)

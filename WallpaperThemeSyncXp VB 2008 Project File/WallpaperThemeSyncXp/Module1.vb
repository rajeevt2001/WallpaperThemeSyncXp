Imports Microsoft.Win32
Imports System.IO
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Diagnostics

Namespace WallpaperThemeSyncXp
    Module ThemeDetector

        <DllImport("user32.dll")> _
        Private Sub keybd_event(ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As UInteger, ByVal dwExtraInfo As UInteger)
        End Sub

        <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
        Private Function SendMessageTimeout(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As UInteger, ByVal lParam As UInteger, ByVal fuFlags As UInteger, ByVal uTimeout As UInteger, ByRef lpdwResult As UInteger) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
        Private Function SendMessageTimeout(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As String, ByVal fuFlags As UInteger, ByVal uTimeout As UInteger, ByRef lpdwResult As IntPtr) As IntPtr
        End Function

        Const HWND_BROADCAST As Integer = &HFFFF
        Const WM_SETTINGCHANGE As Integer = &H1A
        Const SMTO_ABORTIFHUNG As Integer = &H2

        Const SPI_SETDESKWALLPAPER As UInteger = &H14
        Const SPIF_UPDATEINIFILE As UInteger = &H1
        Const SPIF_SENDCHANGE As UInteger = &H2

        Const VK_LWIN As Byte = &H5B
        Const KEYEVENTF_KEYUP As UInteger = &H2

        <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
        Private Function SystemParametersInfo(ByVal uAction As UInteger, ByVal uParam As UInteger, ByVal lpvParam As String, ByVal fuWinIni As UInteger) As Boolean
        End Function

        Const SPI_GETDESKWALLPAPER As UInteger = &H73
        Const MAX_PATH As Integer = 260

        Private screenshot As Bitmap = CaptureScreenWithoutTaskbar()
        Private detecting As Boolean = False

        Structure ColorResult
            Public DominantColor As Color
            Public OriginalColor As Color
        End Structure

        Sub Main()
            Dim currentMsstyle As String = GetCurrentMsstyle()
            Dim currentSubstyle As String = GetCurrentSubstyle()

            If Not IsUserAdministrator() Then
                Console.WriteLine("Program is not run as admin, there may be some errors")
            End If

            If currentMsstyle <> "" Then
                Console.WriteLine("Current Windows and Buttons Style (msstyle): " & currentMsstyle)
                If currentMsstyle.ToLower() = "lunaxp" Then
                    Console.WriteLine("The program can work")
                End If
            Else
                Console.WriteLine("Unable to detect current msstyle.")
            End If

            If currentSubstyle <> "" Then
                Console.WriteLine("Current Substyle: " & currentSubstyle)
            Else
                Console.WriteLine("Unable to detect current substyle.")
            End If

            Console.WriteLine("Minimizing windows to capture desktop...")
            PressWinD()
            Threading.Thread.Sleep(1000)
            screenshot = CaptureScreenWithoutTaskbar()

            Console.WriteLine("Restoring previous windows...")
            PressWinD()

            Dim colorResult As ColorResult = GetDominantColor(screenshot)
            Console.WriteLine("Dominant Screen Color - R: " & colorResult.DominantColor.R & ", G: " & colorResult.DominantColor.G & ", B: " & colorResult.DominantColor.B)
            Dim suggestedColor As String = GetNearestColor(colorResult.DominantColor, colorResult.OriginalColor)
            Console.WriteLine("This color should be applied: " & suggestedColor)

            'Console.WriteLine("Do you want to apply color scheme from the current wallpaper? (Y/N)")
            'Dim userResponse As String = Console.ReadLine().Trim().ToUpper()

            'If userResponse = "Y" Then
            Console.WriteLine("Applying color scheme...")
            ApplyColorScheme(suggestedColor)
            Console.WriteLine("Color applied")
            ' Else
            Console.WriteLine("Exiting program.")
            ' End If
            Threading.Thread.Sleep(1000)
            Environment.Exit(0)
        End Sub

        Function GetCurrentMsstyle() As String
            Try
                Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\ThemeManager")
                    If regKey IsNot Nothing Then
                        Dim themeFilePath As String = regKey.GetValue("DllName")
                        If Not String.IsNullOrEmpty(themeFilePath) Then
                            Return Path.GetFileNameWithoutExtension(themeFilePath)
                        End If
                    End If
                End Using
            Catch ex As Exception
                Console.WriteLine("Error getting current msstyle: " & ex.Message)
            End Try
            Return ""
        End Function

        Function GetCurrentSubstyle() As String
            Try
                Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\ThemeManager")
                    If regKey IsNot Nothing Then
                        Dim substyle As String = regKey.GetValue("ColorName")
                        If Not String.IsNullOrEmpty(substyle) Then
                            Return substyle
                        End If
                    End If
                End Using
            Catch ex As Exception
                Console.WriteLine("Error getting current substyle: " & ex.Message)
            End Try
            Return ""
        End Function

        Function GetDominantColor(ByVal bmp As Bitmap) As ColorResult
            Try
                Dim colorCount As New Dictionary(Of Color, Integer)
                Dim originalColor As Color = Color.Black

                For x As Integer = 0 To bmp.Width - 1 Step 10
                    For y As Integer = 0 To bmp.Height - 1 Step 10
                        Dim pixelColor As Color = bmp.GetPixel(x, y)
                        originalColor = pixelColor
                        If pixelColor.GetSaturation() > 0.1 AndAlso pixelColor.GetBrightness() < 0.9 Then
                            If colorCount.ContainsKey(pixelColor) Then
                                colorCount(pixelColor) += 1
                            Else
                                colorCount(pixelColor) = 1
                            End If
                        End If
                    Next
                Next

                Dim dominantColor As Color = colorCount.OrderByDescending(Function(c) c.Value).FirstOrDefault().Key
                Return New ColorResult With {.DominantColor = dominantColor, .OriginalColor = originalColor}
            Catch ex As Exception
                Console.WriteLine("Error calculating dominant color: " & ex.Message)
                Return New ColorResult With {.DominantColor = Color.Black, .OriginalColor = Color.Black}
            End Try
        End Function


        Function GetNearestColor(ByVal color As Color, ByVal originalColor As Color) As String
            Try
                Dim colors As New Dictionary(Of String, Color)
                colors.Add("Aqua", Drawing.Color.FromArgb(113, 169, 186))
                colors.Add("Black", Drawing.Color.FromArgb(0, 0, 0))
                colors.Add("Blue", Drawing.Color.FromArgb(64, 145, 253))
                colors.Add("Green", Drawing.Color.FromArgb(110, 178, 102))
                colors.Add("Red", Drawing.Color.FromArgb(200, 57, 68))
                colors.Add("Yellow", Drawing.Color.FromArgb(236, 236, 100))
                colors.Add("Orange", Drawing.Color.FromArgb(244, 99, 70))
                colors.Add("Normalcolor", Drawing.Color.FromArgb(144, 109, 209))
                colors.Add("Brown", Drawing.Color.FromArgb(133, 102, 91))

                ' Check for grayscale or near-black/white conditions
                If originalColor.R = originalColor.G AndAlso originalColor.G = originalColor.B Then
                    If originalColor.R >= 60 Then
                        Return "Gray"
                    Else
                        Return "Black"
                    End If
                End If

                Dim nearestColor As String = ""
                Dim minDistance As Double = Double.MaxValue

                For Each kvp In colors
                    Dim dist As Double = Math.Sqrt(Math.Pow(CDbl(color.R) - kvp.Value.R, 2) + _
                                                  Math.Pow(CDbl(color.G) - kvp.Value.G, 2) + _
                                                  Math.Pow(CDbl(color.B) - kvp.Value.B, 2))
                    If dist < minDistance Then
                        minDistance = dist
                        nearestColor = kvp.Key
                    End If
                Next

                Return nearestColor
            Catch ex As Exception
                Console.WriteLine("Error finding nearest color: " & ex.Message)
                Return "Unknown"
            End Try
        End Function

        Sub PressWinD()
            keybd_event(VK_LWIN, 0, 0, 0)
            keybd_event(&H44, 0, 0, 0)
            keybd_event(&H44, 0, KEYEVENTF_KEYUP, 0)
            keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
        End Sub

        Function CaptureScreenWithoutTaskbar() As Bitmap
            Try
                Dim screenBounds As Rectangle = Screen.PrimaryScreen.Bounds
                Dim taskbarHeight As Integer = Screen.PrimaryScreen.Bounds.Height - Screen.PrimaryScreen.WorkingArea.Height

                ' Fallback in case WorkingArea isn’t reliable on XP
                If taskbarHeight <= 0 Then
                    taskbarHeight = 30 ' Typical taskbar height for XP with default settings
                End If

                Dim croppedBounds As New Rectangle(80, 30, screenBounds.Width - 160, screenBounds.Height - taskbarHeight - 80)

                Dim bmp As New Bitmap(croppedBounds.Width, croppedBounds.Height)
                Using g As Graphics = Graphics.FromImage(bmp)
                    g.CopyFromScreen(croppedBounds.Location, Point.Empty, croppedBounds.Size)
                End Using

                Return bmp
            Catch ex As Exception
                Console.WriteLine("Error capturing screen: " & ex.Message)
                Environment.Exit(1)
            End Try
        End Function

        Sub ApplyColorScheme(ByVal colorName As String)
            Try
                Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\ThemeManager", True)
                    If key IsNot Nothing Then
                        ' Check if necessary registry values exist
                        If key.GetValue("DllName") Is Nothing OrElse key.GetValue("ColorName") Is Nothing Then
                            Console.WriteLine("Required registry keys (DllName or ColorName) are missing. Theme application aborted.")
                            Return
                        End If

                        key.SetValue("DllName", "%SystemRoot%\resources\Themes\LunaXP\LunaXP.msstyles")
                        key.SetValue("ColorName", colorName)
                        Console.WriteLine("Updated ColorName to: " & colorName)

                        ' Restart theme-related services
                        RestartService("Themes")

                        ' Restart Explorer to fully apply the theme
                        RestartExplorer()

                        Console.WriteLine("Theme applied after services restart")
                    Else
                        Console.WriteLine("Failed to access registry key.")
                    End If
                End Using
            Catch ex As Exception
                Console.WriteLine("Error applying color scheme: " & ex.Message)
            End Try
        End Sub


        Sub RestartService(ByVal serviceName As String)
            Try
                Process.Start("net", "stop " & serviceName).WaitForExit()
                Process.Start("net", "start " & serviceName).WaitForExit()
                Console.WriteLine(serviceName & " service restarted successfully.")
            Catch ex As Exception
                Console.WriteLine("Error restarting service: " & ex.Message)
            End Try
        End Sub

        Sub RestartExplorer()
            Try
                Dim explorerProcesses = Process.GetProcessesByName("explorer")
                If explorerProcesses.Length > 0 Then
                    For Each proc As Process In explorerProcesses
                        proc.Kill()
                    Next
                    Threading.Thread.Sleep(1000)
                    Process.Start("explorer.exe")
                    Console.WriteLine("Explorer restarted.")
                Else
                    Console.WriteLine("No running Explorer instances found.")
                End If
            Catch ex As Exception
                Console.WriteLine("Error restarting Explorer: " & ex.Message)
            End Try
        End Sub

        Function IsUserAdministrator() As Boolean
            Dim identity = System.Security.Principal.WindowsIdentity.GetCurrent()
            Dim principal = New System.Security.Principal.WindowsPrincipal(identity)
            Return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator)
        End Function


    End Module
End Namespace

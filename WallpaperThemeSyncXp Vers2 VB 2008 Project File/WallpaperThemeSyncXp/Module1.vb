Imports Microsoft.Win32
Imports System.IO
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms

Namespace WallpaperThemeSyncXp
    Module ThemeDetector

        ' --- Configuration Settings ---
        Public Enum ColorCalculationMethod
            NewHueHistogram ' 12-Bin Hue Histogram + HSB Angle Mapping
            OldRgbFrequency ' Exact RGB Frequency + 3D Euclidean Distance
            VibrantWeighted ' Vibrant-Weighted Lower-Region '
        End Enum

        ' Dynamic Settings (Modified via System Tray Menu)
        Private selectedDelayMs As Integer = 5000
        Private selectedMethod As ColorCalculationMethod = ColorCalculationMethod.OldRgbFrequency

        ' --- Win32 Console Window Control ---
        <DllImport("kernel32.dll")> _
        Private Function GetConsoleWindow() As IntPtr
        End Function

        <DllImport("user32.dll")> _
        Private Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)> _
        Private Function AllocConsole() As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)> _
        Private Function FreeConsole() As Boolean
        End Function

        Private Const SW_HIDE As Integer = 0
        Private Const SW_SHOW As Integer = 5

        ' --- Win32 Theme API ---
        <DllImport("uxtheme.dll", EntryPoint:="#65", CharSet:=CharSet.Unicode, SetLastError:=True)> _
        Private Function SetSystemVisualStyle( _
            ByVal pszThemeFileName As String, _
            ByVal pszColorBuff As String, _
            ByVal pszSizeBuff As String, _
            ByVal dwReserved As Integer) As Integer
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto)> _
        Private Function SystemParametersInfo(ByVal uAction As UInteger, ByVal uParam As UInteger, ByVal lpvParam As StringBuilder, ByVal fuWinIni As UInteger) As Boolean
        End Function

        Const SPI_GETDESKWALLPAPER As UInteger = &H73
        Const MAX_PATH As Integer = 260

        ' --- System Tray Components ---
        Private trayIcon As NotifyIcon
        Private trayMenu As ContextMenu
        Private menuToggleConsole As MenuItem

        ' Sub-menu items for dynamic checkmarks
        Private menuMethodHue As MenuItem
        Private menuMethodRgb As MenuItem
        Private menuMethodVibrant As MenuItem

        Private menuDelay05s As MenuItem
        Private menuDelay5s As MenuItem
        Private menuDelay10s As MenuItem
        Private menuDelay20s As MenuItem

        Private consoleVisible As Boolean = False
        Private consoleHandle As IntPtr = IntPtr.Zero

        <STAThread()> _
        Sub Main()
            ' Hide console window immediately on execution start
            consoleHandle = GetConsoleWindow()
            If consoleHandle <> IntPtr.Zero Then
                ShowWindow(consoleHandle, SW_HIDE)
            End If

            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)

            ' Check OS compatibility
            If Not IsWindowsXp() Then
                MessageBox.Show("This application is designed strictly for Windows XP.", _
                                "Wallpaper Theme Sync XP", _
                                MessageBoxButtons.OK, _
                                MessageBoxIcon.Error)
                Return
            End If

            ' Setup System Tray Menu
            InitializeTrayIcon()

            ' Listen for Windows / JBS Wallpaper Change Events
            AddHandler SystemEvents.UserPreferenceChanged, AddressOf OnUserPreferenceChanged

            ' Perform initial sync WITH startup delay
            ThreadPool.QueueUserWorkItem(AddressOf PerformSyncCallback, True)

            ' Keep application thread running in background
            Application.Run()
        End Sub



        Private Sub InitializeTrayIcon()

            trayMenu = New ContextMenu()
            Dim menuSyncNow As New MenuItem("Sync Now", AddressOf OnSyncNowClicked)

            ' --- Color Method Sub-Menu ---
            Dim menuMethod As New MenuItem("Color Detection Method")
            menuMethodHue = New MenuItem("12-Bin Hue Histogram", AddressOf OnMethodHueClicked)
            menuMethodRgb = New MenuItem("RGB Frequency", AddressOf OnMethodRgbClicked)
            menuMethodVibrant = New MenuItem("Vibrant-Weighted", AddressOf OnMethodVibrantClicked)
            menuMethod.MenuItems.Add(menuMethodHue)
            menuMethod.MenuItems.Add(menuMethodRgb)
            menuMethod.MenuItems.Add(menuMethodVibrant)

            ' --- Startup Delay Sub-Menu ---
            Dim menuDelay As New MenuItem("Startup Delay")
            menuDelay05s = New MenuItem("0.5 Second", AddressOf OnDelayClicked)
            menuDelay05s.Tag = 500
            menuDelay5s = New MenuItem("5 Seconds", AddressOf OnDelayClicked)
            menuDelay5s.Tag = 5000
            menuDelay10s = New MenuItem("10 Seconds", AddressOf OnDelayClicked)
            menuDelay10s.Tag = 10000
            menuDelay20s = New MenuItem("20 Seconds", AddressOf OnDelayClicked)
            menuDelay20s.Tag = 20000

            menuDelay.MenuItems.Add(menuDelay05s)
            menuDelay.MenuItems.Add(menuDelay5s)
            menuDelay.MenuItems.Add(menuDelay10s)
            menuDelay.MenuItems.Add(menuDelay20s)

            UpdateMenuCheckmarks()

            menuToggleConsole = New MenuItem("Show Console Log", AddressOf OnToggleConsoleClicked)
            Dim menuExit As New MenuItem("Exit", AddressOf OnExitClicked)

            ' Assembly of Menu Layout
            trayMenu.MenuItems.Add(menuSyncNow)
            trayMenu.MenuItems.Add("-")
            trayMenu.MenuItems.Add(menuMethod)
            trayMenu.MenuItems.Add(menuDelay)
            trayMenu.MenuItems.Add("-")
            trayMenu.MenuItems.Add(menuToggleConsole)
            trayMenu.MenuItems.Add("-")
            trayMenu.MenuItems.Add(menuExit)

            trayIcon = New NotifyIcon()
            trayIcon.Text = "Wallpaper Theme Sync XP"
            trayIcon.Icon = SystemIcons.Application
            trayIcon.ContextMenu = trayMenu
            trayIcon.Visible = True
        End Sub

        Private Sub UpdateMenuCheckmarks()
            ' Update Method Checkmarks
            menuMethodHue.Checked = (selectedMethod = ColorCalculationMethod.NewHueHistogram)
            menuMethodRgb.Checked = (selectedMethod = ColorCalculationMethod.OldRgbFrequency)
            menuMethodVibrant.Checked = (selectedMethod = ColorCalculationMethod.VibrantWeighted)

            ' Update Startup Delay Checkmarks
            menuDelay05s.Checked = (selectedDelayMs = 500)
            menuDelay5s.Checked = (selectedDelayMs = 5000)
            menuDelay10s.Checked = (selectedDelayMs = 10000)
            menuDelay20s.Checked = (selectedDelayMs = 20000)
        End Sub

        ' --- ThreadPool Callback Wrapper ---

        Private Sub PerformSyncCallback(ByVal state As Object)
            Dim applyDelay As Boolean = False
            If state IsNot Nothing AndAlso TypeOf state Is Boolean Then
                applyDelay = CBool(state)
            End If
            PerformSync(applyDelay)
        End Sub

        ' --- Event Handlers ---

        Private Sub OnUserPreferenceChanged(ByVal sender As Object, ByVal e As UserPreferenceChangedEventArgs)
            If e.Category = UserPreferenceCategory.Desktop Then
                LogMessage("Desktop wallpaper change detected (Instant Sync).")
                ' Immediate execution without startup delay when JBS/Windows updates desktop
                ThreadPool.QueueUserWorkItem(AddressOf PerformSyncCallback, False)
            End If
        End Sub

        Private Sub OnSyncNowClicked(ByVal sender As Object, ByVal e As EventArgs)
            LogMessage("Manual sync requested (Instant Sync).")
            ' Immediate execution without startup delay
            ThreadPool.QueueUserWorkItem(AddressOf PerformSyncCallback, False)
        End Sub

        Private Sub OnMethodHueClicked(ByVal sender As Object, ByVal e As EventArgs)
            selectedMethod = ColorCalculationMethod.NewHueHistogram
            UpdateMenuCheckmarks()
            LogMessage("Color detection method set to: 12-Bin Hue Histogram")
        End Sub

        Private Sub OnMethodRgbClicked(ByVal sender As Object, ByVal e As EventArgs)
            selectedMethod = ColorCalculationMethod.OldRgbFrequency
            UpdateMenuCheckmarks()
            LogMessage("Color detection method set to: RGB Frequency")
        End Sub

        Private Sub OnMethodVibrantClicked(ByVal sender As Object, ByVal e As EventArgs)
            selectedMethod = ColorCalculationMethod.VibrantWeighted
            UpdateMenuCheckmarks()
            LogMessage("Color detection method set to: Vibrant-Weighted")
        End Sub

        Private Sub OnDelayClicked(ByVal sender As Object, ByVal e As EventArgs)
            Dim item As MenuItem = TryCast(sender, MenuItem)
            If item IsNot Nothing AndAlso item.Tag IsNot Nothing Then
                selectedDelayMs = CInt(item.Tag)
                UpdateMenuCheckmarks()
                LogMessage("Startup delay set to: " & (selectedDelayMs / 1000.0) & " second(s)")
            End If
        End Sub

        Private Sub OnToggleConsoleClicked(ByVal sender As Object, ByVal e As EventArgs)
            ToggleConsoleWindow(Not consoleVisible)
        End Sub

        Private Sub OnExitClicked(ByVal sender As Object, ByVal e As EventArgs)
            trayIcon.Visible = False
            Application.Exit()
            Environment.Exit(0)
        End Sub

        Private Sub ToggleConsoleWindow(ByVal show As Boolean)
            consoleVisible = show

            If show Then
                consoleHandle = GetConsoleWindow()
                If consoleHandle = IntPtr.Zero Then
                    AllocConsole()
                    consoleHandle = GetConsoleWindow()
                Else
                    ShowWindow(consoleHandle, SW_SHOW)
                End If
                menuToggleConsole.Text = "Hide Console Log"
                LogMessage("Console log initialized.")
            Else
                If consoleHandle <> IntPtr.Zero Then
                    ShowWindow(consoleHandle, SW_HIDE)
                End If
                menuToggleConsole.Text = "Show Console Log"
            End If
        End Sub

        Private Sub LogMessage(ByVal message As String)
            If consoleVisible Then
                Console.WriteLine("[" & DateTime.Now.ToLongTimeString() & "] " & message)
            End If
        End Sub

        ' --- Main Sync Logic ---

        Private Sub PerformSync(ByVal applyDelay As Boolean)
            Try
                If applyDelay AndAlso selectedDelayMs > 0 Then
                    LogMessage(String.Format("Startup delay active. Waiting {0} second(s)...", selectedDelayMs / 1000.0))
                    Thread.Sleep(selectedDelayMs)
                End If

                Dim currentMsstylePath As String = GetCurrentMsstylePath()
                Dim wallpaperPath As String = GetCurrentWallpaperPath()

                If String.IsNullOrEmpty(wallpaperPath) OrElse Not File.Exists(wallpaperPath) Then
                    LogMessage("Error: Wallpaper file not found.")
                    Return
                End If

                LogMessage("Wallpaper Path: " & wallpaperPath)

                Dim dominantColor As Color
                Dim suggestedColor As String

                If selectedMethod = ColorCalculationMethod.VibrantWeighted Then
                    dominantColor = ExtractDominantColor_VibrantWeighted(wallpaperPath)
                    suggestedColor = GetNearestColor_Vibrant(dominantColor)
                ElseIf selectedMethod = ColorCalculationMethod.NewHueHistogram Then
                    dominantColor = ExtractDominantColor_Hue(wallpaperPath)
                    suggestedColor = GetNearestColor_Hue(dominantColor)
                Else
                    dominantColor = ExtractDominantColor_RgbFreq(wallpaperPath)
                    suggestedColor = GetNearestColor_Euclidean(dominantColor)
                End If

                LogMessage(String.Format("Dominant Color - R:{0}, G:{1}, B:{2}", dominantColor.R, dominantColor.G, dominantColor.B))
                LogMessage("Matched Substyle: " & suggestedColor)

                If Not String.IsNullOrEmpty(suggestedColor) AndAlso suggestedColor <> "Unknown" Then
                    ApplyColorScheme(currentMsstylePath, suggestedColor)
                End If

            Catch ex As Exception
                LogMessage("Error during sync: " & ex.Message)
            End Try
        End Sub

        ' --- Theme & Wallpaper Helpers ---

        Function GetCurrentMsstylePath() As String
            Try
                Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\ThemeManager")
                    If regKey IsNot Nothing Then
                        Dim themeFilePath As String = TryCast(regKey.GetValue("DllName"), String)
                        If Not String.IsNullOrEmpty(themeFilePath) Then
                            Return Environment.ExpandEnvironmentVariables(themeFilePath)
                        End If
                    End If
                End Using
            Catch ex As Exception
            End Try
            Return ""
        End Function

        Function GetCurrentWallpaperPath() As String
            Try
                Using regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop")
                    If regKey IsNot Nothing Then
                        Dim wallpaper As String = TryCast(regKey.GetValue("Wallpaper"), String)
                        If Not String.IsNullOrEmpty(wallpaper) AndAlso File.Exists(wallpaper) Then
                            Return wallpaper
                        End If
                    End If
                End Using
            Catch ex As Exception
            End Try

            Try
                Dim sb As New StringBuilder(MAX_PATH)
                If SystemParametersInfo(SPI_GETDESKWALLPAPER, MAX_PATH, sb, 0) Then
                    Dim pathResult As String = sb.ToString()
                    If File.Exists(pathResult) Then Return pathResult
                End If
            Catch ex As Exception
            End Try

            Return ""
        End Function

        ' --- Color Calculations ---

        ' --- New Extraction Function ---
        Function ExtractDominantColor_VibrantWeighted(ByVal imagePath As String) As Color
            Try
                Using originalBmp As New Bitmap(imagePath)
                    ' Sample at a manageable 80x80 resolution for high performance
                    Using smallBmp As New Bitmap(originalBmp, New Size(80, 80))

                        ' 12 Hue Buckets (30 degrees each)
                        Dim bucketWeights(11) As Double
                        Dim bucketR(11) As Double
                        Dim bucketG(11) As Double
                        Dim bucketB(11) As Double

                        Dim totalWeight As Double = 0
                        Dim fallbackR As Long = 0, fallbackG As Long = 0, fallbackB As Long = 0

                        For y As Integer = 0 To smallBmp.Height - 1
                            ' 1. Spatial Weighting: Down-weight top 25% (Sky/Clouds)
                            Dim spatialWeight As Double = If(y < (smallBmp.Height * 0.25), 0.3, 1.0)

                            For x As Integer = 0 To smallBmp.Width - 1
                                Dim pixel As Color = smallBmp.GetPixel(x, y)
                                Dim sat As Single = pixel.GetSaturation()
                                Dim bri As Single = pixel.GetBrightness()

                                fallbackR += pixel.R
                                fallbackG += pixel.G
                                fallbackB += pixel.B

                                ' Filter out extreme darks, lights, and completely gray pixels
                                If sat > 0.12F AndAlso bri > 0.12F AndAlso bri < 0.88F Then

                                    ' 2. Non-linear Saturation Boost (S^2.2) & Luminance Mid-Tone Bias
                                    Dim vibranceWeight As Double = Math.Pow(sat, 2.2) * (1.0 - Math.Abs(bri - 0.5) * 1.2)
                                    If vibranceWeight < 0.01 Then Continue For

                                    Dim combinedWeight As Double = vibranceWeight * spatialWeight

                                    ' 3. Map to 30-degree Hue Bucket
                                    Dim hue As Single = pixel.GetHue()
                                    Dim bucketIndex As Integer = CInt(Math.Floor(hue / 30.0F)) Mod 12

                                    bucketWeights(bucketIndex) += combinedWeight
                                    bucketR(bucketIndex) += pixel.R * combinedWeight
                                    bucketG(bucketIndex) += pixel.G * combinedWeight
                                    bucketB(bucketIndex) += pixel.B * combinedWeight
                                    totalWeight += combinedWeight
                                End If
                            Next
                        Next

                        ' Find the winning hue bucket
                        If totalWeight > 0 Then
                            Dim maxWeight As Double = -1
                            Dim maxIndex As Integer = 0

                            For i As Integer = 0 To 11
                                If bucketWeights(i) > maxWeight Then
                                    maxWeight = bucketWeights(i)
                                    maxIndex = i
                                End If
                            Next

                            Dim w As Double = bucketWeights(maxIndex)
                            If w > 0 Then
                                Return Color.FromArgb(CInt(bucketR(maxIndex) / w), CInt(bucketG(maxIndex) / w), CInt(bucketB(maxIndex) / w))
                            End If
                        End If

                        ' Fallback if image is completely grayscale
                        Dim totalPixels As Integer = smallBmp.Width * smallBmp.Height
                        Return Color.FromArgb(CInt(fallbackR \ totalPixels), CInt(fallbackG \ totalPixels), CInt(fallbackB \ totalPixels))

                    End Using
                End Using
            Catch ex As Exception
                Return Color.Gray
            End Try
        End Function

        ' --- Matching Function with Purple & Green Mapping ---
        Function GetNearestColor_Vibrant(ByVal targetColor As Color) As String
            Try
                Dim sat As Single = targetColor.GetSaturation()
                Dim bri As Single = targetColor.GetBrightness()

                ' Low saturation fallback
                If sat < 0.1F Then
                    If bri < 0.25F Then Return "Black" Else Return "Gray"
                End If

                Dim hue As Single = targetColor.GetHue()

                ' Map Hue ranges directly to target style color names
                Select Case hue
                    Case 0.0F To 15.0F, 345.0F To 360.0F : Return "Red"
                    Case 15.0F To 45.0F : If bri < 0.35F Then Return "Brown" Else Return "Orange"
                    Case 45.0F To 70.0F : Return "Yellow"
                    Case 70.0F To 165.0F : Return "Green"          ' Pic 1 falls here (Trees / River valley)
                    Case 165.0F To 205.0F : Return "Aqua"
                    Case 205.0F To 250.0F : Return "Blue"
                    Case 250.0F To 325.0F : Return "Normalcolor"   ' Pic 2 falls here (Purple / Lavender flowers)
                    Case 325.0F To 345.0F : Return "Red"
                    Case Else : Return "Normalcolor"
                End Select
            Catch ex As Exception
                Return "Unknown"
            End Try
        End Function

        Function ExtractDominantColor_Hue(ByVal imagePath As String) As Color
            Try
                Using originalBmp As New Bitmap(imagePath)
                    Using smallBmp As New Bitmap(originalBmp, New Size(64, 64))
                        Dim hueBuckets(11) As Integer
                        Dim hueR(11) As Long
                        Dim hueG(11) As Long
                        Dim hueB(11) As Long

                        Dim totalVibrantPixels As Integer = 0
                        Dim fallbackR As Long = 0, fallbackG As Long = 0, fallbackB As Long = 0

                        For x As Integer = 0 To smallBmp.Width - 1
                            For y As Integer = 0 To smallBmp.Height - 1
                                Dim pixel As Color = smallBmp.GetPixel(x, y)
                                Dim sat As Single = pixel.GetSaturation()
                                Dim bri As Single = pixel.GetBrightness()

                                fallbackR += pixel.R
                                fallbackG += pixel.G
                                fallbackB += pixel.B

                                If sat > 0.15F AndAlso bri > 0.15F AndAlso bri < 0.85F Then
                                    Dim hue As Single = pixel.GetHue()
                                    Dim bucketIndex As Integer = CInt(Math.Floor(hue / 30.0F)) Mod 12

                                    hueBuckets(bucketIndex) += 1
                                    hueR(bucketIndex) += pixel.R
                                    hueG(bucketIndex) += pixel.G
                                    hueB(bucketIndex) += pixel.B
                                    totalVibrantPixels += 1
                                End If
                            Next
                        Next

                        If totalVibrantPixels > 0 Then
                            Dim maxCount As Integer = -1
                            Dim maxIndex As Integer = 0

                            For i As Integer = 0 To 11
                                If hueBuckets(i) > maxCount Then
                                    maxCount = hueBuckets(i)
                                    maxIndex = i
                                End If
                            Next

                            Dim count As Integer = hueBuckets(maxIndex)
                            Return Color.FromArgb(CInt(hueR(maxIndex) \ count), CInt(hueG(maxIndex) \ count), CInt(hueB(maxIndex) \ count))
                        Else
                            Dim totalPixels As Integer = smallBmp.Width * smallBmp.Height
                            Return Color.FromArgb(CInt(fallbackR \ totalPixels), CInt(fallbackG \ totalPixels), CInt(fallbackB \ totalPixels))
                        End If
                    End Using
                End Using
            Catch ex As Exception
                Return Color.Gray
            End Try
        End Function

        Function GetNearestColor_Hue(ByVal targetColor As Color) As String
            Try
                Dim sat As Single = targetColor.GetSaturation()
                Dim bri As Single = targetColor.GetBrightness()

                If sat < 0.12F Then
                    If bri < 0.25F Then Return "Black" Else Return "Gray"
                End If

                Dim hue As Single = targetColor.GetHue()

                Select Case hue
                    Case 0.0F To 15.0F, 345.0F To 360.0F : Return "Red"
                    Case 15.0F To 45.0F : If bri < 0.4F Then Return "Brown" Else Return "Orange"
                    Case 45.0F To 70.0F : Return "Yellow"
                    Case 70.0F To 165.0F : Return "Green"
                    Case 165.0F To 205.0F : Return "Aqua"
                    Case 205.0F To 260.0F : Return "Blue"
                    Case 260.0F To 315.0F : Return "Normalcolor"
                    Case 315.0F To 345.0F : Return "Red"
                    Case Else : Return "Normalcolor"
                End Select
            Catch ex As Exception
                Return "Unknown"
            End Try
        End Function

        Function ExtractDominantColor_RgbFreq(ByVal imagePath As String) As Color
            Try
                Using bmp As New Bitmap(imagePath)
                    Dim colorCount As New Dictionary(Of Color, Integer)

                    For x As Integer = 0 To bmp.Width - 1 Step 10
                        For y As Integer = 0 To bmp.Height - 1 Step 10
                            Dim pixelColor As Color = bmp.GetPixel(x, y)
                            If pixelColor.GetSaturation() > 0.1F AndAlso pixelColor.GetBrightness() < 0.9F Then
                                If colorCount.ContainsKey(pixelColor) Then
                                    colorCount(pixelColor) += 1
                                Else
                                    colorCount(pixelColor) = 1
                                End If
                            End If
                        Next
                    Next

                    Dim maxCount As Integer = -1
                    Dim dominantColor As Color = Color.Black

                    For Each kvp In colorCount
                        If kvp.Value > maxCount Then
                            maxCount = kvp.Value
                            dominantColor = kvp.Key
                        End If
                    Next

                    Return dominantColor
                End Using
            Catch ex As Exception
                Return Color.Black
            End Try
        End Function

        Function GetNearestColor_Euclidean(ByVal color As Color) As String
            Try
                Dim colors As New Dictionary(Of String, Color)
                colors.Add("Aqua", Color.FromArgb(113, 169, 186))
                colors.Add("Black", Color.FromArgb(0, 0, 0))
                colors.Add("Blue", Color.FromArgb(64, 145, 253))
                colors.Add("Green", Color.FromArgb(110, 178, 102))
                colors.Add("Red", Color.FromArgb(200, 57, 68))
                colors.Add("Yellow", Color.FromArgb(236, 236, 100))
                colors.Add("Orange", Color.FromArgb(244, 99, 70))
                colors.Add("Normalcolor", Color.FromArgb(144, 109, 209))
                colors.Add("Brown", Color.FromArgb(133, 102, 91))

                If color.R = color.G AndAlso color.G = color.B Then
                    If color.R >= 60 Then Return "Gray" Else Return "Black"
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
                Return "Unknown"
            End Try
        End Function

        Sub ApplyColorScheme(ByVal msstylePath As String, ByVal colorName As String)
            Try
                If String.IsNullOrEmpty(msstylePath) OrElse Not File.Exists(msstylePath) Then
                    msstylePath = Environment.ExpandEnvironmentVariables("%SystemRoot%\resources\Themes\LunaXP\LunaXP.msstyles")
                End If

                Dim hr As Integer = SetSystemVisualStyle(msstylePath, colorName, "NormalSize", 0)

                If hr = 0 Then
                    LogMessage("Color scheme applied successfully!")
                    Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\ThemeManager", True)
                        If key IsNot Nothing Then
                            key.SetValue("DllName", msstylePath)
                            key.SetValue("ColorName", colorName)
                        End If
                    End Using
                End If
            Catch ex As Exception
                LogMessage("Error applying color scheme: " & ex.Message)
            End Try
        End Sub

        Function IsWindowsXp() As Boolean
            Dim os As OperatingSystem = Environment.OSVersion
            Return (os.Platform = PlatformID.Win32NT AndAlso os.Version.Major = 5 AndAlso (os.Version.Minor = 1 OrElse os.Version.Minor = 2))
        End Function

    End Module
End Namespace
Imports System.Diagnostics
Imports System.IO
Imports System.Configuration
Imports System.ServiceProcess
Imports System.Text

Public Class ServiceManager
    Implements IDisposable

    Private WithEvents processTimer As System.Timers.Timer
    Private WithEvents configWatcher As FileSystemWatcher
    Private WithEvents logCleanupTimer As System.Timers.Timer
    Private exeProcess As Process = Nothing
    Private config As Dictionary(Of String, String)
    Private configPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "service.ini")
    Private stopTimeout As Integer = 5000 ' 默认停止超时时间为5秒
    Private logWriter As StreamWriter = Nothing
    Private logPath As String = ""
    Private outputBuffer As New StringBuilder()
    Private errorBuffer As New StringBuilder()
    Private restartCount As Integer = 0
    Private lastRestartTime As DateTime = DateTime.MinValue
    Private Const MAX_RESTART_COUNT As Integer = 5
    Private Const RESTART_COUNT_RESET_INTERVAL As Integer = 3600 ' 1小时后重置重启计数
    Private isRunning As Boolean = False
    Private serviceName As String = ""
    Private disposedValue As Boolean

    Private Sub InitializeEventLogging()
        Try
            ' 设置服务名称
            serviceName = "MihomoService"
            
            ' 初始化事件日志源
            If Not EventLog.SourceExists(serviceName) Then
                EventLog.CreateEventSource(serviceName, "Application")
            End If
        Catch ex As Exception
            ' 如果无法创建事件源，至少尝试写入应用程序日志
            Try
                EventLog.WriteEntry("Application", $"无法创建事件日志源: {ex.Message}", EventLogEntryType.Error)
            Catch
                ' 如果连这都失败了，我们真的无能为力了
            End Try
        End Try
    End Sub

    Public Sub New()
        Try
            ' 初始化事件日志源
            InitializeEventLogging()

            ' 确保在最早期就能记录日志
            Try
                EventLog.WriteEntry(serviceName, "服务实例正在初始化...", EventLogEntryType.Information)
            Catch
                ' 忽略事件日志错误
            End Try

            ' 检查配置文件是否存在，不存在则创建默认配置
            If Not File.Exists(configPath) Then
                EventLog.WriteEntry(serviceName, $"配置文件不存在，将在位置创建默认配置: {configPath}", EventLogEntryType.Information)
                CreateDefaultConfig()
            End If

            ' 加载配置
            config = LoadConfigurationFile()
            
            ' 初始化日志
            InitializeLogging()
            
            WriteLog("服务初始化完成")
            
        Catch ex As Exception
            ' 如果在构造函数中发生错误，记录到Windows事件日志
            Try
                EventLog.WriteEntry(serviceName, $"服务初始化失败: {ex.Message}", EventLogEntryType.Error)
                If ex.InnerException IsNot Nothing Then
                    EventLog.WriteEntry(serviceName, $"详细错误: {ex.InnerException.Message}", EventLogEntryType.Error)
                End If
                EventLog.WriteEntry(serviceName, $"错误堆栈: {ex.StackTrace}", EventLogEntryType.Error)
            Catch
                ' 如果连事件日志都写不了，那就真的没办法了
            End Try
            Throw
        End Try
    End Sub

    Public Function CreateWindowsService() As ServiceBase
        Return New CustomService(Me)
    End Function

    ' 用于控制台模式的启动方法
    Public Sub Start()
        Try
            ' 设置定时器查进程状态
            Dim checkInterval As Integer = 5000 ' 默认5
            If config.ContainsKey("checkinterval") AndAlso Integer.TryParse(config("checkinterval"), checkInterval) Then
                ' 使用配置的检查间隔
            End If

            processTimer = New System.Timers.Timer()
            processTimer.Interval = checkInterval
            processTimer.Start()

            StartProcess()
            WriteLog("服务已启动")
            isRunning = True
        Catch ex As Exception
            WriteLog("启动服务时出错: " & ex.Message, EventLogEntryType.Error)
            Throw
        End Try
    End Sub

    ' 用于控制台模式的停止方法
    Public Sub [Stop]()
        Try
            If processTimer IsNot Nothing Then
                processTimer.Stop()
                processTimer.Dispose()
            End If

            StopProcess()
            WriteLog("服务已停止")
            isRunning = False
        Catch ex As Exception
            WriteLog("停止服务时出错: " & ex.Message, EventLogEntryType.Error)
        End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' 不要更改此代码。请将清理代码放入"Dispose(disposing As Boolean)"方法中
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub

    Private Sub WriteLog(message As String, Optional type As EventLogEntryType = EventLogEntryType.Information)
        Try
            ' 控制台输出（仅在交互模式下）
            If Environment.UserInteractive Then
                SyncLock Console.Out
                    Select Case type
                        Case EventLogEntryType.Error
                            Console.ForegroundColor = ConsoleColor.Red
                        Case EventLogEntryType.Warning
                            Console.ForegroundColor = ConsoleColor.Yellow
                        Case Else
                            Console.ForegroundColor = ConsoleColor.Gray
                    End Select

                    Console.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}")
                    Console.ResetColor()
                End SyncLock
            End If

            ' 文件日志（始终写入）
            If logWriter IsNot Nothing Then
                SyncLock logWriter
                    Try
                        logWriter.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{type}] {message}")
                        logWriter.Flush()
                    Catch ex As Exception
                        ' 如果写入文件日志失败，且在服务模式下，记录到事件日志
                        If Not Environment.UserInteractive Then
                            EventLog.WriteEntry(serviceName, $"写入日志文件失败: {ex.Message}", EventLogEntryType.Error)
                        End If
                    End Try
                End SyncLock
            End If

            ' 事件日志（仅在服务模式下记录错误和警告）
            If Not Environment.UserInteractive AndAlso (type = EventLogEntryType.Error OrElse type = EventLogEntryType.Warning) Then
                Try
                    EventLog.WriteEntry(serviceName, message, type)
                Catch ex As Exception
                    ' 如果连事件日志都写不了，那就真的无能为力了
                End Try
            End If

        Catch ex As Exception
            ' 如果所有日志记录方式都失败，且在服务模式下，尝试写入事件日志
            If Not Environment.UserInteractive Then
                Try
                    EventLog.WriteEntry(serviceName, $"日志记录失败: {ex.Message}", EventLogEntryType.Error)
                Catch
                    ' 如果连事件日志都写不了，那就真的无能为力了
                End Try
            End If
        End Try
    End Sub

    Private Sub InitializeConfigWatcher()
        Try
            configWatcher = New FileSystemWatcher()
            configWatcher.Path = Path.GetDirectoryName(configPath)
            configWatcher.Filter = Path.GetFileName(configPath)
            configWatcher.NotifyFilter = NotifyFilters.LastWrite
            configWatcher.EnableRaisingEvents = True
        Catch ex As Exception
            WriteLog($"初始化配置监视器失败: {ex.Message}", EventLogEntryType.Warning)
        End Try
    End Sub

    Private Sub ConfigFileChanged(sender As Object, e As FileSystemEventArgs) Handles configWatcher.Changed
        Try
            ' 等待文件写入完成
            System.Threading.Thread.Sleep(1000)

            ' 重新加载配置
            Dim newConfig = LoadConfigurationFile()
            
            ' 检查关键配置是否改变
            If newConfig.ContainsKey("Program.ExePath") AndAlso newConfig("Program.ExePath") <> config("Program.ExePath") Then
                WriteLog("检测到exe路径变更，需要重启服务才能生效", EventLogEntryType.Warning)
                Return
            End If

            ' 更新配置
            config = newConfig
            WriteLog("配置已重新加载", EventLogEntryType.Information)

            ' 更新日志设置
            UpdateLogging()

        Catch ex As Exception
            WriteLog($"重新加载配置失败: {ex.Message}", EventLogEntryType.Error)
        End Try
    End Sub

    Private Sub InitializeLogging()
        Try
            ' 获取日志目录
            Dim logDir = GetConfigValue("Logging", "LogDir", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs"))
            
            ' 创建日志目录（如果不存在）
            If Not Directory.Exists(logDir) Then
                Directory.CreateDirectory(logDir)
            End If

            ' 设置日志文件名（使用exe名称）
            Dim processName = Path.GetFileNameWithoutExtension(GetConfigValue("Program", "ExePath", ""))
            logPath = Path.Combine(logDir, $"{processName}.log")
            
            ' 关闭现有的日志写入器
            If logWriter IsNot Nothing Then
                Try
                    logWriter.Close()
                    logWriter.Dispose()
                Catch
                    ' 忽略关闭错误
                End Try
            End If

            ' 创建新的日志写入器（覆盖模式）
            logWriter = New StreamWriter(logPath, False, Encoding.UTF8) With {
                .AutoFlush = True
            }
            
            ' 写入启动标记
            logWriter.WriteLine("==========================================")
            logWriter.WriteLine($"服务启动时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            logWriter.WriteLine("==========================================")
            
        Catch ex As Exception
            EventLog.WriteEntry(serviceName, $"初始化日志系统失败: {ex.Message}", EventLogEntryType.Error)
            Throw
        End Try
    End Sub

    Private Sub ProcessOutputDataReceived(sender As Object, e As DataReceivedEventArgs)
        If e.Data IsNot Nothing Then
            Try
                SyncLock outputBuffer
                    outputBuffer.AppendLine(e.Data)
                End SyncLock
                If logWriter IsNot Nothing Then
                    logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss}] {e.Data}")
                End If
            Catch ex As Exception
                WriteLog($"处理程序输出时出错: {ex.Message}", EventLogEntryType.Error)
            End Try
        End If
    End Sub

    Private Sub ProcessErrorDataReceived(sender As Object, e As DataReceivedEventArgs)
        If e.Data IsNot Nothing Then
            Try
                SyncLock errorBuffer
                    errorBuffer.AppendLine(e.Data)
                End SyncLock
                If logWriter IsNot Nothing Then
                    logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss}] [ERROR] {e.Data}")
                End If
            Catch ex As Exception
                WriteLog($"处理程序错误输出时出错: {ex.Message}", EventLogEntryType.Error)
            End Try
        End If
    End Sub

    Private Sub StartProcess()
        Try
            If exeProcess IsNot Nothing Then
                Try
                    If Not exeProcess.HasExited Then
                        WriteLog("进程已在运行中", EventLogEntryType.Warning)
                        Return
                    End If
                    exeProcess.Dispose()
                Catch
                    ' 忽略清理时的错误
                End Try
            End If

            ' 创建新的进程实例
            exeProcess = New Process()
            
            ' 设置进程启动信息
            With exeProcess.StartInfo
                .FileName = GetConfigValue("Program", "ExePath", "")
                .UseShellExecute = False
                .RedirectStandardOutput = True
                .RedirectStandardError = True
                .CreateNoWindow = True
                .StandardOutputEncoding = Encoding.UTF8
                .StandardErrorEncoding = Encoding.UTF8

                ' 设置启动参数
                .Arguments = GetConfigValue("Program", "Arguments", "")

                ' 设置工作目录
                .WorkingDirectory = GetConfigValue("Program", "WorkingDirectory", Path.GetDirectoryName(.FileName))
            End With

            ' 设置事件处理
            AddHandler exeProcess.OutputDataReceived, AddressOf OnProcessOutputDataReceived
            AddHandler exeProcess.ErrorDataReceived, AddressOf OnProcessErrorDataReceived
            AddHandler exeProcess.Exited, AddressOf OnProcessExit
            exeProcess.EnableRaisingEvents = True

            ' 启动进程
            exeProcess.Start()
            WriteLog($"进程已启动，PID: {exeProcess.Id}")

            ' 开始异步读取输出
            exeProcess.BeginOutputReadLine()
            exeProcess.BeginErrorReadLine()

        Catch ex As Exception
            WriteLog($"启动进程失败: {ex.Message}", EventLogEntryType.Error)
            Throw
        End Try
    End Sub

    Private Sub StopProcess()
        Try
            If exeProcess IsNot Nothing AndAlso Not exeProcess.HasExited Then
                ' 尝试优雅关闭
                If GetConfigBool("Process", "GracefulShutdown", True) Then
                    exeProcess.CloseMainWindow()
                    
                    ' 等待进程退出
                    Dim stopTimeout = GetConfigInt("Process", "StopTimeout", 5000)
                    If exeProcess.WaitForExit(stopTimeout) Then
                        If logWriter IsNot Nothing Then
                            logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss}] 进程已正常退出")
                        End If
                        Return
                    End If
                End If

                ' 如果优雅关闭失败或配置为不使用优雅关闭，则强制结束进程
                If logWriter IsNot Nothing Then
                    logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss}] 正在强制结束进程")
                End If
                exeProcess.Kill()
            End If
            
        Catch ex As Exception
            EventLog.WriteEntry(serviceName, $"停止进程时出错: {ex.Message}", EventLogEntryType.Error)
        Finally
            If logWriter IsNot Nothing Then
                Try
                    logWriter.WriteLine("==========================================")
                    logWriter.WriteLine($"服务停止时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                    logWriter.WriteLine("==========================================")
                    logWriter.Close()
                    logWriter = Nothing
                Catch
                    ' 忽略关闭日志时的错误
                End Try
            End If
        End Try
    End Sub

    Private Function LoadConfigurationFile() As Dictionary(Of String, String)
        Try
            EventLog.WriteEntry(serviceName, $"正在加载配置文件: {configPath}", EventLogEntryType.Information)
            
            Dim newConfig = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim currentSection As String = ""

            ' 读取配置文件
            Dim lines = File.ReadAllLines(configPath, Encoding.UTF8)
            EventLog.WriteEntry(serviceName, $"成功读取配置文件，共 {lines.Length} 行", EventLogEntryType.Information)

            For Each line In lines
                ' 跳过空行和注释
                If String.IsNullOrWhiteSpace(line) OrElse line.TrimStart().StartsWith(";") Then
                    Continue For
                End If

                ' 检查是否是节名
                If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                    currentSection = line.Substring(1, line.Length - 2)
                    Continue For
                End If

                ' 解析配置项
                Dim parts = line.Split(New Char() {"="c}, 2)
                If parts.Length = 2 Then
                    Dim key = parts(0).Trim()
                    Dim value = parts(1).Trim()
                    
                    ' 处理值中的引号
                    value = ParseConfigValue(value)
                    
                    ' 使用节名作为前缀，以避免不同节中的同名配置项冲突
                    If Not String.IsNullOrEmpty(currentSection) Then
                        key = currentSection & "." & key
                    End If
                    
                    newConfig(key) = value
                    EventLog.WriteEntry("Application", $"读取配置项: {key} = {value}", EventLogEntryType.Information)
                End If
            Next

            ' 验证要的配置
            If Not newConfig.ContainsKey("Program.ExePath") Then
                EventLog.WriteEntry("Application", "配置文件中缺少必要的 'Program.ExePath' 配置项", EventLogEntryType.Error)
                Throw New Exception("配置文件中缺少必要的 'Program.ExePath' 配置项")
            End If

            ' 验证可执行文件是否存在
            Dim exePath = newConfig("Program.ExePath")
            EventLog.WriteEntry("Application", $"正在检查可执行文件: {exePath}", EventLogEntryType.Information)
            
            If Not File.Exists(exePath) Then
                EventLog.WriteEntry("Application", $"指定的可执行文件不存在: {exePath}", EventLogEntryType.Error)
                Throw New FileNotFoundException($"指定的可执行文件不存在: {exePath}", exePath)
            End If

            Return newConfig

        Catch ex As Exception
            EventLog.WriteEntry("Application", $"加载配置文件失败: {ex.Message}", EventLogEntryType.Error)
            Throw
        End Try
    End Function

    Private Function ParseConfigValue(value As String) As String
        Try
            value = value.Trim()
            
            ' 处理整个值的引号
            If value.StartsWith("""") AndAlso value.EndsWith("""") Then
                value = value.Substring(1, value.Length - 2)
            End If
            
            ' 处理值中的引号转义
            If value.Contains("""") Then
                ' 将连续的两个引号替换为单个引号
                value = value.Replace("""""", """")
                
                ' 处理参数中的引号，例如 --config "config.json"
                Dim parts = value.Split(" "c)
                For i = 0 To parts.Length - 1
                    If parts(i).StartsWith("""") AndAlso parts(i).EndsWith("""") Then
                        parts(i) = parts(i).Substring(1, parts(i).Length - 2)
                    End If
                Next
                value = String.Join(" ", parts)
            End If
            
            Return value
        Catch ex As Exception
            EventLog.WriteEntry("Application", $"解析配置值时出错: {value}, {ex.Message}", EventLogEntryType.Warning)
            Return value
        End Try
    End Function

    ' 辅助方法：获取配置值，如果不存在则返回默认值
    Private Function GetConfigValue(section As String, key As String, defaultValue As String) As String
        Dim fullKey = section & "." & key
        If config.ContainsKey(fullKey) Then
            Return config(fullKey)
        End If
        Return defaultValue
    End Function

    ' 辅助方法：获取配置的整数值，如果不存在或无效则返回默认值
    Private Function GetConfigInt(section As String, key As String, defaultValue As Integer) As Integer
        Dim value = GetConfigValue(section, key, defaultValue.ToString())
        Dim result As Integer
        If Integer.TryParse(value, result) Then
            Return result
        End If
        Return defaultValue
    End Function

    ' 辅助方法：获取配置的布尔值，如果不存在或无效则返回默认值
    Private Function GetConfigBool(section As String, key As String, defaultValue As Boolean) As Boolean
        Dim value = GetConfigValue(section, key, defaultValue.ToString())
        Dim result As Boolean
        If Boolean.TryParse(value, result) Then
            Return result
        End If
        Return defaultValue
    End Function

    Private Sub LoadConfiguration()
        Try
            config = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            ' 如果配置文件不存在，创建默认配置文件
            If Not File.Exists(configPath) Then
                CreateDefaultConfig()
                WriteLog("配置文件不存在，已创建默认配置文件。请编辑配置文件后重启服务。", EventLogEntryType.Warning)
                Throw New Exception("请编辑配置文件后重启服务")
            End If

            config = LoadConfigurationFile()

        Catch ex As Exception
            WriteLog("加载配置失败: " & ex.Message, EventLogEntryType.Error)
            Throw
        End Try
    End Sub

    Private Sub UpdateLogging()
        Try
            ' 关闭现有的日志文件
            If logWriter IsNot Nothing Then
                logWriter.WriteLine("==========================================")
                logWriter.WriteLine($"日志文件闭时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                logWriter.WriteLine("==========================================")
                logWriter.Close()
                logWriter = Nothing
            End If

            ' 重新初始化日志
            InitializeLogging()

        Catch ex As Exception
            EventLog.WriteEntry(serviceName, $"更新日志设置失败: {ex.Message}", EventLogEntryType.Error)
        End Try
    End Sub

    Private Sub ProcessTimer_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs) Handles processTimer.Elapsed
        Try
            MonitorProcessHealth()
        Catch ex As Exception
            WriteLog($"进程监控定时器出错: {ex.Message}", EventLogEntryType.Error)
        End Try
    End Sub

    Private Function CanRestartProcess() As Boolean
        ' 检查是否超过最大重启次数
        If DateTime.Now - lastRestartTime > TimeSpan.FromSeconds(RESTART_COUNT_RESET_INTERVAL) Then
            ' 重置计数
            restartCount = 0
            lastRestartTime = DateTime.Now
        End If

        restartCount += 1
        If restartCount > MAX_RESTART_COUNT Then
            EventLog.WriteEntry(serviceName, $"进程重启次数过多（{restartCount}次），停止重启", EventLogEntryType.Error)
            Return False
        End If

        Return True
    End Function

    Private Sub MonitorProcessHealth()
        Try
            If exeProcess Is Nothing OrElse exeProcess.HasExited Then
                Return
            End If

            ' 检查内存使用
            Dim memoryUsageMB = exeProcess.WorkingSet64 / (1024 * 1024)
            Dim maxMemoryMB = GetConfigInt("Process", "MaxMemory", 500)

            If memoryUsageMB > maxMemoryMB Then
                If logWriter IsNot Nothing Then
                    Try
                        SyncLock logWriter
                            logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss}] 内存使用超过限制: {memoryUsageMB:N0}MB > {maxMemoryMB}MB，将重启进程")
                            logWriter.Flush()
                        End SyncLock
                    Catch
                        ' 忽略日志写入错误
                    End Try
                End If
                RestartProcess()
            End If

        Catch ex As Exception
            EventLog.WriteEntry(serviceName, $"监控进程健康状态时出错: {ex.Message}", EventLogEntryType.Error)
        End Try
    End Sub

    Private Sub RestartProcess()
        Try
            If Not CanRestartProcess() Then Return

            StopProcess()
            System.Threading.Thread.Sleep(1000) ' 等待1秒
            StartProcess()

            If logWriter IsNot Nothing Then
                Try
                    SyncLock logWriter
                        logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss}] 进程已重启，当前重启次数: {restartCount}")
                        logWriter.Flush()
                    End SyncLock
                Catch
                    ' 忽略日志写入错误
                End Try
            End If

        Catch ex As Exception
            EventLog.WriteEntry(serviceName, $"重启进程失败: {ex.Message}", EventLogEntryType.Error)
        End Try
    End Sub

    Private Sub CreateDefaultConfig()
        Try
            EventLog.WriteEntry(serviceName, "开始创建默认配置文件...", EventLogEntryType.Information)
            
            ' 创建默认配置内容
            Dim defaultConfig As New StringBuilder()
            defaultConfig.AppendLine("; Windows 服务配置文件")
            defaultConfig.AppendLine("; 本配置文件用于控制Windows服务如何管理目标程序")
            defaultConfig.AppendLine("; 配置文件使用UTF-8编码，支持中文")
            defaultConfig.AppendLine()
            
            defaultConfig.AppendLine("[Program]")
            defaultConfig.AppendLine("; 要运行的程序完整路径（必需配置）")
            defaultConfig.AppendLine("; 如果路径包含空格，必须使用双引号")
            defaultConfig.AppendLine(";ExePath=""C:\Program Files\MyApp\MyApp.exe""")
            defaultConfig.AppendLine()
            defaultConfig.AppendLine("; 启动参数（可选）")
            defaultConfig.AppendLine(";Arguments=--config ""config.json""")
            defaultConfig.AppendLine()
            defaultConfig.AppendLine("; 工作目录（可选，默认使用exe所在目录）")
            defaultConfig.AppendLine(";WorkingDirectory=C:\Program Files\MyApp")
            defaultConfig.AppendLine()
            
            defaultConfig.AppendLine("[Logging]")
            defaultConfig.AppendLine("; 日志文件保存目录（可选，默认为服务目录下的logs文件夹）")
            defaultConfig.AppendLine(";LogDir=C:\Logs\MyApp")
            defaultConfig.AppendLine()
            
            defaultConfig.AppendLine("[Process]")
            defaultConfig.AppendLine("; 检查进程状态的时间间隔，单位：毫秒（可选，默认5000）")
            defaultConfig.AppendLine(";CheckInterval=5000")
            defaultConfig.AppendLine()
            defaultConfig.AppendLine("; 最大内存使用限制，单位：MB（可选，默认500）")
            defaultConfig.AppendLine(";MaxMemory=500")
            defaultConfig.AppendLine()
            defaultConfig.AppendLine("; 停止进程的超时时间，单位：毫秒（可选，默认5000）")
            defaultConfig.AppendLine(";StopTimeout=5000")
            defaultConfig.AppendLine()
            defaultConfig.AppendLine("; 是否使用优雅关闭（可选，默认true）")
            defaultConfig.AppendLine(";GracefulShutdown=true")

            ' 创建配置文件目录（如果不存在）
            Dim configDir = Path.GetDirectoryName(configPath)
            If Not Directory.Exists(configDir) Then
                Directory.CreateDirectory(configDir)
                EventLog.WriteEntry(serviceName, $"创建配置文件目录: {configDir}", EventLogEntryType.Information)
            End If

            ' 写入配置文件
            File.WriteAllText(configPath, defaultConfig.ToString(), Encoding.UTF8)
            EventLog.WriteEntry(serviceName, $"已创建默认配置文件: {configPath}", EventLogEntryType.Information)

        Catch ex As Exception
            EventLog.WriteEntry(serviceName, $"创建默认配置文件失败: {ex.Message}", EventLogEntryType.Error)
            Throw
        End Try
    End Sub

    Private Class CustomService
        Inherits ServiceBase

        Private ReadOnly manager As ServiceManager

        Public Sub New(manager As ServiceManager)
            Try
                Me.manager = manager
                Me.ServiceName = manager.serviceName
                Me.CanStop = True
                Me.CanPauseAndContinue = False
                Me.AutoLog = True

                ' 记录服务创建
                EventLog.WriteEntry(serviceName, "Windows服务实例正在创建...", EventLogEntryType.Information)
            Catch ex As Exception
                EventLog.WriteEntry(serviceName, $"创建Windows服务实例失败: {ex.Message}", EventLogEntryType.Error)
                Throw
            End Try
        End Sub

        Protected Overrides Sub OnStart(ByVal args() As String)
            Try
                EventLog.WriteEntry(serviceName, "Windows服务OnStart被调用...", EventLogEntryType.Information)
                manager.Start()
            Catch ex As Exception
                EventLog.WriteEntry(serviceName, $"Windows服务启动失败: {ex.Message}", EventLogEntryType.Error)
                If ex.InnerException IsNot Nothing Then
                    EventLog.WriteEntry(serviceName, $"详细错误: {ex.InnerException.Message}", EventLogEntryType.Error)
                End If
                EventLog.WriteEntry(serviceName, $"错误堆栈: {ex.StackTrace}", EventLogEntryType.Error)
                Throw
            End Try
        End Sub

        Protected Overrides Sub OnStop()
            Try
                EventLog.WriteEntry(serviceName, "Windows服务OnStop被调用...", EventLogEntryType.Information)
                manager.Stop()
            Catch ex As Exception
                EventLog.WriteEntry(serviceName, $"Windows服务停止失败: {ex.Message}", EventLogEntryType.Error)
                Throw
            End Try
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            Try
                If disposing Then
                    manager.Dispose()
                End If
                MyBase.Dispose(disposing)
            Catch ex As Exception
                EventLog.WriteEntry(serviceName, $"Windows服务释放资源失败: {ex.Message}", EventLogEntryType.Error)
                Throw
            End Try
        End Sub
    End Class

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' 释放托管资源
                If processTimer IsNot Nothing Then
                    processTimer.Dispose()
                End If
                If configWatcher IsNot Nothing Then
                    configWatcher.Dispose()
                End If
                If logCleanupTimer IsNot Nothing Then
                    logCleanupTimer.Dispose()
                End If
                If logWriter IsNot Nothing Then
                    logWriter.Dispose()
                End If
                If exeProcess IsNot Nothing Then
                    Try
                        If Not exeProcess.HasExited Then
                            exeProcess.Kill()
                        End If
                        exeProcess.Dispose()
                    Catch
                        ' 忽略清理时错误
                    End Try
                End If
            End If

            ' 释放未托管资源
            disposedValue = True
        End If
    End Sub

    Private Sub OnProcessExit(sender As Object, e As EventArgs)
        Try
            If logWriter IsNot Nothing Then
                Try
                    SyncLock logWriter
                        logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss}] 进程已退出")
                        logWriter.Flush()
                    End SyncLock
                Catch
                    ' 忽略日志写入错误
                End Try
            End If
            
            ' 检查是否超过最大重启次数
            If DateTime.Now - lastRestartTime > TimeSpan.FromSeconds(RESTART_COUNT_RESET_INTERVAL) Then
                ' 重置计数
                restartCount = 0
                lastRestartTime = DateTime.Now
            End If
            
            restartCount += 1
            If restartCount > MAX_RESTART_COUNT Then
                EventLog.WriteEntry(serviceName, $"进程重启次数过多（{restartCount}次），停止重启", EventLogEntryType.Error)
                Return
            End If
            
            ' 重新启动进程
            StartProcess()
            
        Catch ex As Exception
            ' 这是服务级别的错误，应该记录到事件日志
            EventLog.WriteEntry(serviceName, $"处理进程退出事件时出错: {ex.Message}", EventLogEntryType.Error)
        End Try
    End Sub

    Private Sub OnProcessOutputDataReceived(sender As Object, e As DataReceivedEventArgs)
        If Not String.IsNullOrEmpty(e.Data) Then
            ' 只写入文件日志
            If logWriter IsNot Nothing Then
                Try
                    SyncLock logWriter
                        logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss}] [OUTPUT] {e.Data}")
                        logWriter.Flush()
                    End SyncLock
                Catch
                    ' 忽略日志写入错误
                End Try
            End If
        End If
    End Sub

    Private Sub OnProcessErrorDataReceived(sender As Object, e As DataReceivedEventArgs)
        If Not String.IsNullOrEmpty(e.Data) Then
            ' 只写入文件日志
            If logWriter IsNot Nothing Then
                Try
                    SyncLock logWriter
                        logWriter.WriteLine($"[{DateTime.Now:HH:mm:ss}] [ERROR] {e.Data}")
                        logWriter.Flush()
                    End SyncLock
                Catch
                    ' 忽略日志写入错误
                End Try
            End If
        End If
    End Sub
End Class
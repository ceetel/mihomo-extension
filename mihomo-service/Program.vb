Imports System.ServiceProcess

Module Program
    Sub Main(args As String())
        Try
            Using manager As New ServiceManager()
                If Environment.UserInteractive Then
                    ' 控制台模式
                    Console.WriteLine("服务启动中...")
                    Console.WriteLine("按 Q 键停止服务")

                    ' 启动服务
                    manager.Start()

                    ' 等待用户输入
                    While True
                        Dim key = Console.ReadKey(True)
                        If key.Key = ConsoleKey.Q Then
                            Exit While
                        End If
                    End While

                    ' 停止服务
                    manager.Stop()
                    Console.WriteLine("服务已停止")
                Else
                    ' Windows服务模式
                    Using service = manager.CreateWindowsService()
                        ServiceBase.Run(service)
                    End Using
                End If
            End Using
        Catch ex As Exception
            If Environment.UserInteractive Then
                Console.WriteLine($"错误: {ex.Message}")
                Console.WriteLine("按任意键退出...")
                Console.ReadKey(True)
            Else
                EventLog.WriteEntry("Service Startup", "服务启动失败: " & ex.Message, EventLogEntryType.Error)
            End If
        End Try
    End Sub
End Module 
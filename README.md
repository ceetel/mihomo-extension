# 一个通用的Windows服务
这个项目本意是用来启动Mihomo(原Clash-Meta)内核, 因为在网上没有找到一个通用的可以启动现有commandline exe文件的Windows服务, 所以迫不得已肝了这么小东西, 后来思来想去不如直接写一个更通用的Windows服务, 用来启动commandline软件, 这也就是这个服务诞生时的样子.

## 简要的使用方法
打开程序本体, 他会自动创建服务配置文件, 修改完配置文件后, 使用sc命令管理服务
## 配置文件(以clash-meta举例):
尽管启动参数是可选配置项, 但是由于Windows的service默认使用的用户是LocalSystem, 而非用户登陆的账户, 所以一般必须要配置绝对路径的启动参数. 然而我并没有找到一种可靠的方法使用登陆的账户来启动服务, 所以说个人认为对于系统服务的配置上Windows完全比不上Linux的Systemd来的简单.


```ini
; Windows 服务配置文件
; 本配置文件用于控制Windows服务如何管理目标程序
; 配置文件使用UTF-8编码，支持中文

[Program]
; 要运行的程序完整路径（必需配置）
; 如果路径包含空格，必须使用双引号
ExePath="C:\Users\soong\AppData\Local\Microsoft\WinGet\Packages\MetaCubeX.mihomo_Microsoft.Winget.Source_8wekyb3d8bbwe\mihomo-windows-amd64.exe"

; 启动参数（可选）
Arguments=-d "C:\Users\soong\.config\mihomo"

; 工作目录（可选，默认使用exe所在目录）
;WorkingDirectory=C:\Program Files\MyApp

[Logging]
; 日志文件保存目录（可选，默认为服务目录下的logs文件夹）
;LogDir=C:\Logs\MyApp

[Process]
; 检查进程状态的时间间隔，单位：毫秒（可选，默认5000）
;CheckInterval=5000

; 最大内存使用限制，单位：MB（可选，默认500）
;MaxMemory=500

; 停止进程的超时时间，单位：毫秒（可选，默认5000）
;StopTimeout=5000

; 是否使用优雅关闭（可选，默认true）
;GracefulShutdown=true

```
## 创建, 删除, 管理服务
使用[SC](https://learn.microsoft.com/zh-cn/windows-server/administration/windows-commands/sc-config)创建和管理

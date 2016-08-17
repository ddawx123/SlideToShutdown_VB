Attribute VB_Name = "Module"
'Xiaoding Studio Copyright 2016
'就不做更多定时关机、重启、瞬间断电的功能了，类似的VB模块均已发送到吾爱破解论坛共享，要集成的可以到我的另外几篇帖子下载模块并在本程序中实现参数调用。

Option Explicit

Public errnumber  '定义了一个错误编码存储专用的变量

Sub Main()
If App.PrevInstance = True Then  '多开检测
MsgBox "本程序正在运行，不允许多次执行！请打开任务管理器检查是否有残余进程，如果存在，则应关闭它，然后再次尝试启动本程序。", vbOKOnly + vbSystemModal + vbExclamation, "多开检测模块"
Else


'参数调用检测


If Command() = "" Then

'以下是主程序

On Error GoTo errControl  '如果找不到下面将要调用的接口文件，则转到错误处理流程。
Shell Environ("windir") & "\sysnative\SlideToShutDown.exe"  '调用系统接口执行滑动关机。
errControl:  '找不到文件的错误处理。
errnumber = Err.Number  '错误编码导出到自定义变量
If errnumber = "53" Then  '检测，如果为53号，则说明文件缺失，系统自动弹窗警告。
MsgBox "亲，你的系统不是win8或win10，本程序调用的是系统自带接口，由于您的系统缺失这个接口，程序无法继续工作！请升级或降级系统后重试。", vbCritical + vbSystemModal + vbOKOnly, "找不到接口"
Else
'什么都不做
End If

'主程序结束

Else
MsgBox "参数调用模块没有添加，请自行修改源码并重新编译到本程序！", vbOKOnly + vbSystemModal, "找不到模块"
End If

'参数调用检测结束

End If
End Sub

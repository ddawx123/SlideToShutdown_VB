Attribute VB_Name = "Module"
'Xiaoding Studio Copyright 2016
'�Ͳ������ඨʱ�ػ���������˲��ϵ�Ĺ����ˣ����Ƶ�VBģ����ѷ��͵��ᰮ�ƽ���̳����Ҫ���ɵĿ��Ե��ҵ����⼸ƪ��������ģ�鲢�ڱ�������ʵ�ֲ������á�

Option Explicit

Public errnumber  '������һ���������洢ר�õı���

Sub Main()
If App.PrevInstance = True Then  '�࿪���
MsgBox "�������������У���������ִ�У�����������������Ƿ��в�����̣�������ڣ���Ӧ�ر�����Ȼ���ٴγ�������������", vbOKOnly + vbSystemModal + vbExclamation, "�࿪���ģ��"
Else


'�������ü��


If Command() = "" Then

'������������

On Error GoTo errControl  '����Ҳ������潫Ҫ���õĽӿ��ļ�����ת�����������̡�
Shell Environ("windir") & "\sysnative\SlideToShutDown.exe"  '����ϵͳ�ӿ�ִ�л����ػ���
errControl:  '�Ҳ����ļ��Ĵ�����
errnumber = Err.Number  '������뵼�����Զ������
If errnumber = "53" Then  '��⣬���Ϊ53�ţ���˵���ļ�ȱʧ��ϵͳ�Զ��������档
MsgBox "�ף����ϵͳ����win8��win10����������õ���ϵͳ�Դ��ӿڣ���������ϵͳȱʧ����ӿڣ������޷������������������򽵼�ϵͳ�����ԡ�", vbCritical + vbSystemModal + vbOKOnly, "�Ҳ����ӿ�"
Else
'ʲô������
End If

'���������

Else
MsgBox "��������ģ��û����ӣ��������޸�Դ�벢���±��뵽������", vbOKOnly + vbSystemModal, "�Ҳ���ģ��"
End If

'�������ü�����

End If
End Sub

var WSHShell = WScript.CreateObject("WScript.Shell");

 WSHShell.Popup("������� ������");
 ////////////////////////////////////////////////////////////////////////////////////////////
 // ��������� ���������
 // ���������� ���
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\debug_log", "D:\\CopyVKDBLog\\DebugLog.log");
 // ���-���� ������ ���������.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\txt_log", "D:\\CopyVKDBLog\\MyLogFile.log");
 // ������� �������� �������� �����.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\logsrc", "\\\\MAINSERT\\D$\\VKDB\\archive");
 // ������� � ������� �� �������� �������� ����.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\logdest", "C:\\arc_vkdb");

 ////////////////////////////////////////////////////////////////////////////////////////////// 
 // Oracle - ����
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\orahost", "vkdb");
 // ������������
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\orauser", "sys");
 // ������
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\orapwd", "man");

 /////////////////////////////////////////////////////////////////////////////////////////////
 // �������� ���������
 // smtp-������ ����� ������� �������� ���������.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\smtp", "192.168.1.115");
 // �� ����� ���� ���� ���������.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\mail_from", "prg@protek34.tomsknet.ru");
 // ���� ���� ���������.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\mail_to", "treeclimber@mail.ru");






var WSHShell = WScript.CreateObject("WScript.Shell");

 WSHShell.Popup("Создаем раздел");
 ////////////////////////////////////////////////////////////////////////////////////////////
 // настройка каталогов
 // отладочный лог
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\debug_log", "D:\\CopyVKDBLog\\DebugLog.log");
 // лог-файл работы программы.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\txt_log", "D:\\CopyVKDBLog\\MyLogFile.log");
 // каталог источник архивных логов.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\logsrc", "\\\\MAINSERT\\D$\\VKDB\\archive");
 // каталог в который мы копируем архивные логи.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\logdest", "C:\\arc_vkdb");

 ////////////////////////////////////////////////////////////////////////////////////////////// 
 // Oracle - хост
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\orahost", "vkdb");
 // пользователь
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\orauser", "sys");
 // пароль
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\orapwd", "man");

 /////////////////////////////////////////////////////////////////////////////////////////////
 // почтовые настройки
 // smtp-сервер через который отсылаем сообщения.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\smtp", "192.168.1.115");
 // от имени кого шлем сообщения.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\mail_from", "prg@protek34.tomsknet.ru");
 // кому шлем сообщения.
 WSHShell.RegWrite("HKLM\\SOFTWARE\\Copy_Log1\\mail_to", "treeclimber@mail.ru");






////////////////////////////////////////////////////////////////////////////////////////
///          JScript для копирования архивных логов. 
/// Скрипт также контролирует непрерывности архивных логов Oracle.
///                 v. 3.0. от 15.12.2006.
///////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////
// Создаем объекты 
///////////////////////////////////////////////////////////////////////////////////////
// чтение настроек
var WSHShell = WScript.CreateObject("WScript.Shell");

// OCX модуль для отправки почтового сообщения.
var  postie = WScript.CreateObject("Postie.Postie");

// для работы с файлами
var fso = WScript.CreateObject("Scripting.FileSystemObject");

// Работа с Oracle через OO4O
var OraSession = WScript.CreateObject("OracleInProcServer.XOraSession");

/////////////////////////////////////////////////////////////////////////////////////////
// Инициализируем параметры.
/////////////////////////////////////////////////////////////////////////////////////////
/// источник и приёмник для архив-логов
/// сюда пишет архив-логи Oracle
var pathSrc   = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\logsrc");
/// сюда мы их копируем 
var pathDst   = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\logdest");

// файлы для записи лог-файлов для работы данной програмки
var Mylog    = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\txt_log");
var DebugLog = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\debug_log");

// параметры для коннекта к Oracle
var OraLogin = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\orauser");
var OraPwd   = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\orapwd");
var OraHost  = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\orahost");

// параметры e-mail
var parSMTP = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\smtp");
var parFROM = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\mail_from");
var parTO   = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\mail_to");
///////////////////////////////////////////////////////////////////////////////////////////



//////////////////////////////////////////////////////////////////////////////////////////////////
// Инициализируем файловые ресурсы.
//////////////////////////////////////////////////////////////////////////////////////////////////
// Открываем или создаем отладочный лог-файл
if (fso.FileExists(DebugLog))
     var txtDebugOut = fso.OpenTextFile(DebugLog,8,true); // Если файл есть - открыть для добавления
  else
     var txtDebugOut = fso.OpenTextFile(DebugLog,2,true); // Если файла нет - открыть для записи

// Открываем или создаем текстовый лог-файл
  if (fso.FileExists(Mylog))
     var txtStreamOut = fso.OpenTextFile(Mylog,8,true); // Если файл есть - открыть для добавления
  else
     var txtStreamOut = fso.OpenTextFile(Mylog,2,true); // Если файла нет - открыть для записи

// открываем для чтения файл с sql-скриптом.
var txtStreamIn = fso.OpenTextFile("log_hist.sql");
var txtStreamDel = fso.OpenTextFile("log_hist_del.sql");
///////////////////////////////////////////////////////////////////////////////////////////////////

// читаем текст запроса в переменную.
var txtSQL = " ";
while( !txtStreamIn.atEndOfStream ) {
      txtSQL = txtSQL+" "+txtStreamIn.ReadLine();
} 

var txtSQLdel = " ";

while( !txtStreamDel.atEndOfStream ) {
      txtSQLdel = txtSQLdel+" "+txtStreamDel.ReadLine();
} 


//WScript.Echo(txtSQL);


////////////////////////////////////////////////////////////////////////////////////////////
//  Попытки соединения с серверами
////////////////////////////////////////////////////////////////////////////////////////////
// коннектимся к SMTP-серверу
try {
  postie.ConnectSmtp(parSMTP, 25);
}
catch(e) {
   if (e != 0) {
    txtDebugOut.WriteLine(new Date().toLocaleString()+" SMTP-server error : " +e.message);
    postie.Disconnect();
  }
}

// коннектимся к серверу-Oracle
try {
  var OraDatabase = OraSession.OpenDatabase(OraHost, OraLogin+"/"+OraPwd, 0);
  // запрос к БД 
  var  lSQL = OraDatabase.CreateDynaset(txtSQL, 0); 
  var  dSQL = OraDatabase.CreateDynaset(txtSQLdel, 0); 

  OraSession.Open;
  lSQL.Execute;
  dSQL.Execute;
}
catch(e) {
   if (e != 0) {
    txtDebugOut.WriteLine(new Date().toLocaleString()+" Oracle error : " +e.message);
    postie.Post(parFROM, parTO, "Oracle error!", new Date().toLocaleString()+" "+e.message, 0, 1);
    OraSession.Close; 
    WScript.Quit(1);
  }
}


txtDebugOut.WriteLine(new Date().toLocaleString()+" copy_log.js run.");

///название архив-лога
var logFile = "";

///////////////////////////////////////////////////////////////////////////////////////////////////
// читаем фетч запроса
  while( !lSQL.EOF ) {
// прочитали строчку запроса в переменную
    logFile = lSQL.Fields(0).Value; 

//  WScript.Echo(pathSrc+"\\"+logFile);
// проверяем наличие исходного файла.
    if (fso.FileExists(pathSrc+"\\"+logFile)) {
      var oFile = fso.GetFile(pathSrc+"\\"+logFile);
// проверяем есть-ли уже в каталоге-приёмнике такой файл.
      if ( !fso.FileExists(pathDst+"\\"+logFile) ) {
// если еще нет - то копируем
         try {
           oFile.Copy (pathDst+"\\"+logFile, false);
         }
         catch(e) {
           if (e != 0) {
             txtStreamOut.WriteLine(new Date().toLocaleString()+"Error copy File "+logFile+" "+e.message);
             postie.Post(parFROM, parTO, "Error copy!!!",new Date().toLocaleString()+"Error copy File "+logFile+" "+e.message, 0, 1);
           }
         }    
         txtStreamOut.WriteLine(new Date().toLocaleString()+" "+logFile+" Ok!");         
      }
    }
    else {
      txtStreamOut.WriteLine(new Date().toLocaleString()+" "+"Error ARCH-log "+pathSrc+"\\"+logFile+" not exist!");
      postie.Post(parFROM, parTO, "Error ARCH!!!", new Date().toLocaleString()+" "+"Error ARCH-log "+pathSrc+"\\"+logFile+" not exist!", 0, 1);
    }
    logFile = "";
    lSQL.MoveNext();
  }

// читаем фетч второго запроса для удаления архив-логов
  while( !dSQL.EOF ) {
// прочитали строчку запроса в переменную
    logFile = dSQL.Fields(0).Value; 

//  WScript.Echo(pathSrc+"\\"+logFile);
// проверяем наличие исходного файла.
    if (fso.FileExists(pathSrc+"\\"+logFile)) {
      var oFile = fso.GetFile(pathSrc+"\\"+logFile);
// проверяем есть-ли уже в каталоге-приёмнике такой файл.
      if ( fso.FileExists(pathDst+"\\"+logFile) ) {
// если есть в приемнике  - то удалим из источника
         try {
           oFile.Delete();
           txtStreamOut.WriteLine(new Date().toLocaleString()+" "+"Delete source File "+logFile);
         }
         catch(e) {
           if (e != 0) {
             txtStreamOut.WriteLine(new Date().toLocaleString()+"Error delete source File "+logFile+" "+e.message);
             postie.Post(parFROM, parTO, "Error delete source!!!",new Date().toLocaleString()+"Error copy File "+logFile+" "+e.message, 0, 1);
           }
         }    
         txtStreamOut.WriteLine(new Date().toLocaleString()+" "+logFile+" Ok!");         
      }     
    }   
// а теперь удалим из приемника    
    if (fso.FileExists(pathDst+"\\"+logFile)) {
       var dFile = fso.GetFile(pathDst+"\\"+logFile);
       try {
         dFile.Delete();
         txtStreamOut.WriteLine(new Date().toLocaleString()+" "+"Delete destination File "+logFile);
       }
       catch(e) {
         if (e != 0) {
            txtStreamOut.WriteLine(new Date().toLocaleString()+"Error delete destination File "+logFile+" "+e.message);
            postie.Post(parFROM, parTO, "Error delete destination!!!",new Date().toLocaleString()+"Error copy File "+logFile+" "+e.message, 0, 1);
         }
       }
    }

    logFile = "";
    dSQL.MoveNext();
  }



  // postie.Post(parFROM, parTO, "Ok!", new Date().toLocaleString()+" ALL Rights!", 0, 1);
  OraSession.Close; 
  postie.Disconnect();
///////////////////////////////////////////////////////////////////////////////////////////////////



////////////////////////////////////////////////////////////////////////////////////////
///          JScript ��� ����������� �������� �����. 
/// ������ ����� ������������ ������������� �������� ����� Oracle.
///                 v. 3.0. �� 15.12.2006.
///////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////
// ������� ������� 
///////////////////////////////////////////////////////////////////////////////////////
// ������ ��������
var WSHShell = WScript.CreateObject("WScript.Shell");

// OCX ������ ��� �������� ��������� ���������.
var  postie = WScript.CreateObject("Postie.Postie");

// ��� ������ � �������
var fso = WScript.CreateObject("Scripting.FileSystemObject");

// ������ � Oracle ����� OO4O
var OraSession = WScript.CreateObject("OracleInProcServer.XOraSession");

/////////////////////////////////////////////////////////////////////////////////////////
// �������������� ���������.
/////////////////////////////////////////////////////////////////////////////////////////
/// �������� � ������� ��� �����-�����
/// ���� ����� �����-���� Oracle
var pathSrc   = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\logsrc");
/// ���� �� �� �������� 
var pathDst   = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\logdest");

// ����� ��� ������ ���-������ ��� ������ ������ ���������
var Mylog    = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\txt_log");
var DebugLog = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\debug_log");

// ��������� ��� �������� � Oracle
var OraLogin = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\orauser");
var OraPwd   = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\orapwd");
var OraHost  = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\orahost");

// ��������� e-mail
var parSMTP = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\smtp");
var parFROM = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\mail_from");
var parTO   = WSHShell.RegRead("HKLM\\SOFTWARE\\Copy_Log1\\mail_to");
///////////////////////////////////////////////////////////////////////////////////////////



//////////////////////////////////////////////////////////////////////////////////////////////////
// �������������� �������� �������.
//////////////////////////////////////////////////////////////////////////////////////////////////
// ��������� ��� ������� ���������� ���-����
if (fso.FileExists(DebugLog))
     var txtDebugOut = fso.OpenTextFile(DebugLog,8,true); // ���� ���� ���� - ������� ��� ����������
  else
     var txtDebugOut = fso.OpenTextFile(DebugLog,2,true); // ���� ����� ��� - ������� ��� ������

// ��������� ��� ������� ��������� ���-����
  if (fso.FileExists(Mylog))
     var txtStreamOut = fso.OpenTextFile(Mylog,8,true); // ���� ���� ���� - ������� ��� ����������
  else
     var txtStreamOut = fso.OpenTextFile(Mylog,2,true); // ���� ����� ��� - ������� ��� ������

// ��������� ��� ������ ���� � sql-��������.
var txtStreamIn = fso.OpenTextFile("log_hist.sql");
var txtStreamDel = fso.OpenTextFile("log_hist_del.sql");
///////////////////////////////////////////////////////////////////////////////////////////////////

// ������ ����� ������� � ����������.
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
//  ������� ���������� � ���������
////////////////////////////////////////////////////////////////////////////////////////////
// ����������� � SMTP-�������
try {
  postie.ConnectSmtp(parSMTP, 25);
}
catch(e) {
   if (e != 0) {
    txtDebugOut.WriteLine(new Date().toLocaleString()+" SMTP-server error : " +e.message);
    postie.Disconnect();
  }
}

// ����������� � �������-Oracle
try {
  var OraDatabase = OraSession.OpenDatabase(OraHost, OraLogin+"/"+OraPwd, 0);
  // ������ � �� 
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

///�������� �����-����
var logFile = "";

///////////////////////////////////////////////////////////////////////////////////////////////////
// ������ ���� �������
  while( !lSQL.EOF ) {
// ��������� ������� ������� � ����������
    logFile = lSQL.Fields(0).Value; 

//  WScript.Echo(pathSrc+"\\"+logFile);
// ��������� ������� ��������� �����.
    if (fso.FileExists(pathSrc+"\\"+logFile)) {
      var oFile = fso.GetFile(pathSrc+"\\"+logFile);
// ��������� ����-�� ��� � ��������-�������� ����� ����.
      if ( !fso.FileExists(pathDst+"\\"+logFile) ) {
// ���� ��� ��� - �� ��������
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

// ������ ���� ������� ������� ��� �������� �����-�����
  while( !dSQL.EOF ) {
// ��������� ������� ������� � ����������
    logFile = dSQL.Fields(0).Value; 

//  WScript.Echo(pathSrc+"\\"+logFile);
// ��������� ������� ��������� �����.
    if (fso.FileExists(pathSrc+"\\"+logFile)) {
      var oFile = fso.GetFile(pathSrc+"\\"+logFile);
// ��������� ����-�� ��� � ��������-�������� ����� ����.
      if ( fso.FileExists(pathDst+"\\"+logFile) ) {
// ���� ���� � ���������  - �� ������ �� ���������
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
// � ������ ������ �� ���������    
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



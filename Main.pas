unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, IdHTTP, IdURI, IdSSLOpenSSL,
  Vcl.ComCtrls, DAScript, UniScript, OraCall, Data.DB, DBAccess, Ora, OraScript,
  Vcl.ShellAnimations, ShellApi, MemDS, IdIOHandler,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, Registry, System.IOUtils, cxGraphics,
  cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxContainer, cxEdit,
  cxLabel, System.Zip, TlHelp32, PsAPI;

const
   WM_INSIGHT_UPDATE = WM_USER + 1;
   NullDate = -700000;


type
  TfrmMain = class(TForm)
    Animate1: TAnimate;
    Script: TOraScript;
    OraSession: TOraSession;
    ShellResources1: TShellResources;
    IdHTTP: TIdHTTP;
    IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;
    qryTmp: TOraQuery;
    lblConnection: TcxLabel;
    lblScript: TcxLabel;
    procedure FormShow(Sender: TObject);
    procedure ScriptError(Sender: TObject; E: Exception; SQL: string;
      var Action: TErrorAction);
    procedure ScriptBeforeExecute(Sender: TObject; var SQL: string;
          var Omit: Boolean);
  private
    { Private procedure }

    procedure OnInsightUpdate(var Msg: TMessage); message WM_INSIGHT_UPDATE;

    function DownloadFile(const Url: string; TargetFileName,
                                  User, Pass: string): boolean;

    function ProcessExists(anExeFileName: string): Boolean;

  public
    { Public declarations }

  end;

var
  frmMain: TfrmMain;

   function TzSpecificLocalTimeToSystemTime(lpTimeZoneInformation: PTimeZoneInformation; var lpLocalTime, lpUniversalTime: TSystemTime): BOOL; stdcall;
   function SystemTimeToTzSpecificLocalTime(lpTimeZoneInformation: PTimeZoneInformation; var lpUniversalTime,lpLocalTime: TSystemTime): BOOL; stdcall;

implementation

   function TzSpecificLocalTimeToSystemTime; external kernel32 name 'TzSpecificLocalTimeToSystemTime';
   function SystemTimeToTzSpecificLocalTime; external kernel32 name 'SystemTimeToTzSpecificLocalTime';


{$R *.dfm}


Function DateTime2UnivDateTime(d:TDateTime):TDateTime;
var
 TZI:TTimeZoneInformation;
 LocalTime, UniversalTime:TSystemTime;
begin
  GetTimeZoneInformation(tzi);
  DateTimeToSystemTime(d,LocalTime);
  TzSpecificLocalTimeToSystemTime(@tzi,LocalTime,UniversalTime);
  Result := SystemTimeToDateTime(UniversalTime);
end;

Function UnivDateTime2LocalDateTime(d:TDateTime):TDateTime;
var
 TZI:TTimeZoneInformation;
 LocalTime, UniversalTime:TSystemTime;
begin
  GetTimeZoneInformation(tzi);
  DateTimeToSystemTime(d,UniversalTime);
  SystemTimeToTzSpecificLocalTime(@tzi,UniversalTime,LocalTime);
  Result := SystemTimeToDateTime(LocalTime);
end;

function TfrmMain.DownloadFile(const Url: string; TargetFileName,
                                  User, Pass: string): boolean;
var
//   IdHTTP: TIdHTTP;
   Response,
   FullTargetFileName,
   IndexedFullTargetFileName,
   tempFile,
   NewDocPath,
   NewDocName,
   AParsedDocName: string;
   FileStream: TFileStream;
//   LHandler: TIdSSLIOHandlerSocketOpenSSL;
   FileHandle: NativeInt;
   numBytes: integer;
   URI: TIdURI;
   bFileError: boolean;
begin
   Result := False;
   bFileError := False;
   if (Url <> '') then
   begin
      try
//         IdHTTP := TIdHTTP.Create(nil);
//         LHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
//         LHandler.SSLOptions.Method := sslvTLSv1;
         try
            FullTargetFileName := TargetFileName;

            If FileExists(FullTargetFileName) = False then
            begin
               FileHandle := NativeInt(FileCreate(FullTargetFileName));
               if (FileHandle = -1) then
               begin
                  bFileError := True;
               end;
               FileClose(FileHandle);
            end;

            if bFileError = False then
            begin
               IdHTTP.AllowCookies := True;
               IdHTTP.HandleRedirects := True;

               IdHTTP.Request.Username := User;
               IdHTTP.Request.Password := Pass;
               IdHTTP.Request.BasicAuthentication := False;
               IdHTTP.HTTPOptions := [hoInProcessAuth];

               IdHTTP.Request.ContentType := 'application/x-www-form-urlencoded';
               IdHTTP.Request.Connection := 'keep-alive';
               // Download file
               try
//                  IdHTTP.IOHandler:=LHandler;
                  URI := TIdURI.Create(Url);
                  URI.Username := User;
                  URI.Password := Pass;

                  FileStream := TFileStream.Create(FullTargetFileName, fmOpenReadWrite);
                  try
                     IdHTTP.Get(URI.GetFullURI([ofAuthInfo]), FileStream);
                     numBytes := IdHTTP.response.contentLength;
                  finally
                     FileStream.Free;
                  end;
               finally
//                  LHandler.Free;
                  Result := True;
                  IdHTTP.Disconnect;
               end;
            end;

         except
            on E: Exception do
               ShowMessage(E.Message);
         end;
      finally
//         IdHTTP.Free;
         URI.Free;
      end;
   end
   else
      Result := True;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
   Animate1.Active := True;
   Application.ProcessMessages;
   PostMessage(self.Handle, WM_INSIGHT_UPDATE, 0, 0);
end;

procedure TfrmMain.OnInsightUpdate(var Msg: TMessage);
var
   dbName,
   SourceString: string;
   WebfileDate,
   LocalscriptDate,
   UnivFileDate,
   UnivLocalDate:  TDateTime;
   ContentLength: integer;
   LocalPath,
   sDirectConn,
   sOldInsight,
   sRenamedInsight,
   sDownloadInsight,
   sOldExesDir,
   sRenamedInsightFile,
   sSQLFile,
   sZipFile: string;
   reg: TRegistry;
   AZipfile: TZipFile;
begin
   LocalPath := IncludeTrailingBackslash(ExtractFileDir(Application.ExeName)) + 'Insight.exe';
   if ProcessExists(LocalPath) = False then
   begin
      Application.ProcessMessages;
      sOldExesDir := IncludeTrailingBackslash(ExtractFileDir(LocalPath))+ 'old_executable';

      // create variable to store renamed insight
      // we'll test to see if we can rename file before we do anything
      sOldInsight := LocalPath;
      sRenamedInsightFile := StringReplace('insight_'+ DatetimeToStr(Now) + '.exe', '/', '_', [rfReplaceAll, rfIgnoreCase]);
      sRenamedInsightFile := StringReplace(sRenamedInsightFile , ':', '_', [rfReplaceAll, rfIgnoreCase]);
      sRenamedInsightFile := StringReplace(sRenamedInsightFile , 'PM', '', [rfReplaceAll, rfIgnoreCase]);
      sRenamedInsightFile := StringReplace(sRenamedInsightFile , 'AM', '', [rfReplaceAll, rfIgnoreCase]);
      sRenamedInsightFile := StringReplace(sRenamedInsightFile , ' ', '', [rfReplaceAll, rfIgnoreCase]);

      Application.ProcessMessages;
      if TDirectory.Exists(sOldExesDir) = False then
         TDirectory.CreateDirectory(sOldExesDir);

      Application.ProcessMessages;
      if RenameFile(sOldInsight, sRenamedInsightFile) = True then
      begin
         sRenamedInsight := IncludeTrailingBackslash(ExtractFileDir(LocalPath)) + sRenamedInsightFile;

         sDownloadInsight := IncludeTrailingBackslash(ExtractFileDir(LocalPath)) + 'insight_new.exe';

         try
            if TFile.Exists(IncludeTrailingBackslash(sOldExesDir) + sRenamedInsightFile) = False then
               TFile.Move(sRenamedInsight, IncludeTrailingBackslash(sOldExesDir) + sRenamedInsightFile)
            else
               TFile.Delete(IncludeTrailingBackslash(ExtractFileDir(LocalPath)) + sRenamedInsightFile);
         except
           //
         end;

         Application.ProcessMessages;
         LocalPath := '';
         if TFile.Exists('ChilkatDelphi32.dll') = False then
         begin
            SourceString := 'http://releases.bhlinsight.com/FileReleases/Files/ChilkatDelphi32.zip';
            sZipFile := IncludeTrailingBackslash(ExtractFileDir(Application.ExeName)) + 'ChilkatDelphi32.zip';
            Application.ProcessMessages;
            DownloadFile(SourceString, 'ChilkatDelphi32.zip', 'InsightFileDownload' ,'regdeL99!');
            Application.ProcessMessages;
            AZipFile := TZipfile.Create;
            try
               AZipfile.Open(sZipFile, zmRead);
               AZipfile.ExtractAll('');
               AZipfile.Close;
            finally
               AZipfile.Free;
               TFile.delete(sZipFile);
            end;
         end;

         try
            SourceString := 'http://releases.bhlinsight.com/FileReleases/Files/latest.txt';
            contentLength:=0;
            try
               Idhttp.Head(SourceString);
               WebFileDate:= idhttp.response.LastModified;
               ContentLength := idhttp.response.ContentLength;
            except end;
         except
//
         end;

         Application.ProcessMessages;
         DownloadFile('http://releases.bhlinsight.com/FileReleases/Files/Insight.zip',
                         'insight_new.zip', 'InsightFileDownload' ,'regdeL99!');
         Application.ProcessMessages;
         DownloadFile('http://releases.bhlinsight.com/FileReleases/Files/latest.txt',
                         'latest.sql', 'InsightFileDownload' ,'regdeL99!');

         Application.ProcessMessages;

         try
            reg := TRegistry.Create;
            try
               reg.RootKey := HKEY_CURRENT_USER;
               if reg.OpenKey('Software\Colateral\Axiom\Database', False) then
               begin
                  dbName := reg.ReadString('Server Name');
                  sDirectConn := reg.ReadString('Net');
                  reg.CloseKey;
               end;
            finally
               reg.Free;
            end;

            with OraSession do
            begin
               Options.Direct := (sDirectConn = 'Y');

               Server := dbName;
               Username := 'axiom';
               Password := 'regdeL99';
               try
                  Connect;
               except
                  try
                     Username := 'axiom';
                     Password := 'axiom';
                     Connect;
                  except
                     Application.Terminate;
                  end;
               end;
            end;
         finally
            try
               Script.SQL.Text := 'ALTER TABLE AXIOM.SYSTEMFILE ADD (SCRIPT_EXEC_DATE  DATE)';
               Script.Execute;
            except

            end;

            with qryTmp do
            begin
               Close;
               SQL.Text := 'select SCRIPT_EXEC_DATE as script_exec_date from systemfile';
               Open;
               LocalscriptDate := FieldByName('script_exec_date').AsDateTime;
               Close;
            end;

            UnivFileDate := DateTime2UnivDateTime(WebFileDate);
            UnivLocalDate := DateTime2UnivDateTime(LocalscriptDate);

            if ((UnivLocalDate < UnivFileDate) or (LocalscriptDate = NullDate)) then
            begin
               try
                  Application.ProcessMessages;
                  if (OraSession.Connected = True) then
                  begin
                     Application.ProcessMessages;
                     with qryTmp do
                     begin
                        Close;
                        SQL.Text := 'update systemfile set SCRIPT_EXEC_DATE = SYSDATE';
                        ExecSQL;
                        Close;
                     end;
                     Application.ProcessMessages;

                     sSQLFile := IncludeTrailingBackslash(ExtractFileDir(Application.ExeName)) + 'latest.sql';
                     lblConnection.Caption := 'Connected to: ' + OraSession.Server;
                     lblScript.Caption := sSQLFile;
                     lblConnection.Visible := True;
                     lblScript.Visible := True;
                     Application.ProcessMessages;
                     Script.SQL.LoadFromFile(sSQLFile);
//                     Script.Delimiter := '/';
                     Script.Execute;

                     Application.ProcessMessages;

                     Orasession.Disconnect;
                  end
                  else
                     ShowMessage('Database update script did not run.  Please contact I.T. support or BHL Insight.');
               finally
                  TFile.Delete('latest.sql');
               end;
            end
            else
               TFile.delete('latest.sql');

            try
 //              LocalPath := ExtractFileDir(Application.ExeName);

               Application.ProcessMessages;

{               sOldInsight := IncludeTrailingBackslash(LocalPath) + 'insight.exe';
               sRenamedInsightFile := StringReplace('insight_'+ DatetimeToStr(Now) + '.exe', '/', '_', [rfReplaceAll, rfIgnoreCase]);
               sRenamedInsightFile := StringReplace(sRenamedInsightFile , ':', '_', [rfReplaceAll, rfIgnoreCase]);
               sRenamedInsightFile := StringReplace(sRenamedInsightFile , 'PM', '', [rfReplaceAll, rfIgnoreCase]);
               sRenamedInsightFile := StringReplace(sRenamedInsightFile , 'AM', '', [rfReplaceAll, rfIgnoreCase]);
               sRenamedInsightFile := StringReplace(sRenamedInsightFile , ' ', '_', [rfReplaceAll, rfIgnoreCase]);  }

 //              sRenamedInsight := IncludeTrailingBackslash(LocalPath) + sRenamedInsightFile;

//               sDownloadInsight := IncludeTrailingBackslash(LocalPath) + 'insight_new.exe';

//               sOldExesDir := IncludeTrailingBackslash(LocalPath)+ 'old_executable';

{               if TDirectory.Exists(sOldExesDir) = False then
                  TDirectory.CreateDirectory(sOldExesDir);
               RenameFile(sOldInsight, sRenamedInsight);

               try
                  if TFile.Exists(IncludeTrailingBackslash(sOldExesDir) + sRenamedInsightFile) = False then
                     TFile.Move(sRenamedInsight, IncludeTrailingBackslash(sOldExesDir) + sRenamedInsightFile)
                  else
                     TFile.Delete(IncludeTrailingBackslash(LocalPath) + sRenamedInsightFile);
               except
                 //
               end;  }

               sZipFile := IncludeTrailingBackslash(ExtractFileDir(Application.ExeName)) + 'insight_new.zip';
               AZipFile := TZipfile.Create;
               try
                  AZipfile.Open(sZipFile, zmRead);
                  AZipfile.ExtractAll('');
                  AZipfile.Close;
               finally
                  AZipfile.Free;
                  TFile.delete(sZipFile);
               end;

               Application.ProcessMessages;

               RenameFile(sDownloadInsight, sOldInsight);
               Application.ProcessMessages;
               Animate1.Active := False;
               ShellExecute(Self.Handle,'open', 'insight.exe', nil, nil, SW_SHOWNORMAL);
            finally
               Application.Terminate;
            end;
         end;
      end
      else
      begin
         MessageDlg('Cannot rename Insight.  The Update process cannot continue. '+chr(13)+'Make sure that all instances of Insight are shutdown and try again.',
                 mtError, [mbOk], 0, mbOk);
         Application.Terminate;
      end;
   end
   else
   begin
      MessageDlg('Insight is still running.  The Update process cannot continue. '+chr(13)+'Make sure that all instances of Insight are shutdown and try again.',
                 mtError, [mbOk], 0, mbOk);
      Application.Terminate;
   end;
end;

procedure TfrmMain.ScriptBeforeExecute(Sender: TObject; var SQL: string;
  var Omit: Boolean);
begin
   Application.ProcessMessages;
end;

procedure TfrmMain.ScriptError(Sender: TObject; E: Exception; SQL: string;
  var Action: TErrorAction);
begin
   Action := eaContinue;
end;

function TfrmMain.ProcessExists(anExeFileName: string): Boolean;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
  fullPath: string;
  myHandle: THandle;
  myPID: DWORD;
begin
  // wsyma 2016-04-20 Erkennung, ob ein Prozess in einem bestimmten Pfad schon gestartet wurde.
  // Detection wether a process in a certain path is allready started.
  // http://stackoverflow.com/questions/876224/how-to-check-if-a-process-is-running-using-delphi
  // http://swissdelphicenter.ch/en/showcode.php?id=2010
  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
  Result := False;
  while Integer(ContinueLoop) <> 0 do
  begin
    if UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase(ExtractFileName(anExeFileName)) then
    begin
      myPID := FProcessEntry32.th32ProcessID;
      myHandle := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, False, myPID);
      if myHandle <> 0 then
      try
        SetLength(fullPath, MAX_PATH);
        if GetModuleFileNameEx(myHandle, 0, PChar(fullPath), MAX_PATH) > 0 then
        begin
          SetLength(fullPath, StrLen(PChar(fullPath)));
          if UpperCase(fullPath) = UpperCase(anExeFileName) then
            Result := True;
        end else
          fullPath := '';
      finally
        CloseHandle(myHandle);
      end;
      if Result then
        Break;
    end;
    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
  CloseHandle(FSnapshotHandle);
end;

end.

object frmMain: TfrmMain
  Left = 0
  Top = 0
  ActiveControl = Animate1
  BorderStyle = bsDialog
  Caption = 'Insight Update'
  ClientHeight = 122
  ClientWidth = 312
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Animate1: TAnimate
    Left = 16
    Top = 16
    Width = 272
    Height = 60
    CommonAVI = aviCopyFile
    DoubleBuffered = True
    ParentDoubleBuffered = False
    StopFrame = 22
  end
  object lblConnection: TcxLabel
    Left = 16
    Top = 82
    AutoSize = False
    Caption = 'Connection'
    Properties.ShowAccelChar = False
    Properties.ShowEndEllipsis = True
    Transparent = True
    Visible = False
    Height = 17
    Width = 272
  end
  object lblScript: TcxLabel
    Left = 16
    Top = 102
    AutoSize = False
    Caption = 'Script'
    Properties.ShowEndEllipsis = True
    Transparent = True
    Visible = False
    Height = 17
    Width = 272
  end
  object Script: TOraScript
    BeforeExecute = ScriptBeforeExecute
    OnError = ScriptError
    Session = OraSession
    Left = 128
    Top = 48
  end
  object OraSession: TOraSession
    Options.Direct = True
    Options.IPVersion = ivIPBoth
    Server = 'localhost:1521:insight'
    Left = 72
    Top = 8
  end
  object ShellResources1: TShellResources
    Left = 264
    Top = 56
  end
  object IdHTTP: TIdHTTP
    IOHandler = IdSSLIOHandlerSocketOpenSSL1
    AllowCookies = True
    ProxyParams.BasicAuthentication = False
    ProxyParams.ProxyPort = 0
    Request.ContentLength = -1
    Request.ContentRangeEnd = -1
    Request.ContentRangeStart = -1
    Request.ContentRangeInstanceLength = -1
    Request.Accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
    Request.BasicAuthentication = False
    Request.UserAgent = 'Mozilla/3.0 (compatible; Indy Library)'
    Request.Ranges.Units = 'bytes'
    Request.Ranges = <>
    HTTPOptions = [hoForceEncodeParams]
    Left = 232
  end
  object IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL
    MaxLineAction = maException
    Port = 0
    DefaultPort = 0
    SSLOptions.Mode = sslmUnassigned
    SSLOptions.VerifyMode = []
    SSLOptions.VerifyDepth = 0
    Left = 272
    Top = 8
  end
  object qryTmp: TOraQuery
    Session = OraSession
    Left = 8
    Top = 32
  end
end

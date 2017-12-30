unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  LYTray, Menus, StdCtrls, Buttons, ADODB,
  ActnList, AppEvnts, ComCtrls, ToolWin, ExtCtrls,
  registry,inifiles,Dialogs,
  StrUtils, DB,ComObj,Variants,Math;

type
  TfrmMain = class(TForm)
    LYTray1: TLYTray;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    ApplicationEvents1: TApplicationEvents;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ActionList1: TActionList;
    editpass: TAction;
    about: TAction;
    stop: TAction;
    ToolButton2: TToolButton;
    ToolButton5: TToolButton;
    ToolButton9: TToolButton;
    OpenDialog1: TOpenDialog;
    ADOConn_BS: TADOConnection;
    BitBtn3: TBitBtn;
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure ApplicationEvents1Activate(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
  private
    { Private declarations }
    procedure WMSyscommand(var message:TWMMouse);message WM_SYSCOMMAND;
    procedure UpdateConfig;{配置文件生效}
    function LoadInputPassDll:boolean;
    function MakeDBConn:boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction;

const
  CR=#$D+#$A;
  STX=#$2;ETX=#$3;ACK=#$6;NAK=#$15;
  sCryptSeed='lc';//加解密种子
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='错误!请与开发商联系!' ;
  IniSection='Setup';

var
  ConnectString:string;
  GroupName:string;//
  SpecType:string ;//
  SpecStatus:string ;//
  CombinID:string;//
  LisFormCaption:string;//
  QuaContSpecNoG:string;
  QuaContSpecNo:string;
  QuaContSpecNoD:string;
  EquipChar:string;
  MrConnStr:string;
  ifConnSucc:boolean;
  ifRecLog:boolean;//是否记录调试日志

  hnd:integer;
  bRegister:boolean;

  orderid,sampleid,patientid,acqutime,picturepath:string;
  
{$R *.dfm}

function ifRegister:boolean;
var
  HDSn,RegisterNum,EnHDSn:string;
  configini:tinifile;
  pEnHDSn:Pchar;
begin
  result:=false;
  
  HDSn:=GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'');

  CONFIGINI:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  RegisterNum:=CONFIGINI.ReadString(IniSection,'RegisterNum','');
  CONFIGINI.Free;
  pEnHDSn:=EnCryptStr(Pchar(HDSn),sCryptSeed);
  EnHDSn:=StrPas(pEnHDSn);

  if Uppercase(EnHDSn)=Uppercase(RegisterNum) then result:=true;

  if not result then messagedlg('对不起,您没有注册或注册码错误,请注册!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//是否集成登录模式

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('连接数据库', '服务器', '');
  initialcatalog := Ini.ReadString('连接数据库', '数据库', '');
  ifIntegrated:=ini.ReadBool('连接数据库','集成登录模式',false);
  userid := Ini.ReadString('连接数据库', '用户', '');
  password := Ini.ReadString('连接数据库', '口令', '107DFC967CDCFAAF');
  Ini.Free;
  //======解密password
  pInStr:=pchar(password);
  pDeStr:=DeCryptStr(pInStr,sCryptSeed);
  setlength(password,length(pDeStr));
  for i :=1  to length(pDeStr) do password[i]:=pDeStr[i-1];
  //==========

  result := result + 'user id=' + UserID + ';';
  result := result + 'password=' + Password + ';';
  result := result + 'data source=' + datasource + ';';
  result := result + 'Initial Catalog=' + initialcatalog + ';';
  result := result + 'provider=' + 'SQLOLEDB.1' + ';';
  //Persist Security Info,表示ADO在数据库连接成功后是否保存密码信息
  //ADO缺省为True,ADO.net缺省为False
  //程序中会传ADOConnection信息给TADOLYQuery,故设置为True
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ctext        :string;
  reg          :tregistry;
begin
  ConnectString:=GetConnectString;
  
  UpdateConfig;
  DateTimePicker1.DateTime:=now;
  if ifRegister then bRegister:=true else bRegister:=false;  

  lytray1.Hint:='数据接收服务'+ExtractFileName(Application.ExeName);

//=============================初始化密码=====================================//
    reg:=tregistry.Create;
    reg.RootKey:=HKEY_CURRENT_USER;
    reg.OpenKey('\sunyear',true);
    ctext:=reg.ReadString('pass');
    if ctext='' then
    begin
        reg:=tregistry.Create;
        reg.RootKey:=HKEY_CURRENT_USER;
        reg.OpenKey('\sunyear',true);
        reg.WriteString('pass','JIHONM{');
        //MessageBox(application.Handle,pchar('感谢您使用智能监控系统，'+chr(13)+'请记住初始化密码：'+'lc'),
        //            '系统提示',MB_OK+MB_ICONinformation);     //WARNING
    end;
    reg.CloseKey;
    reg.Free;
//============================================================================//
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
    if not LoadInputPassDll then exit;
    application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  show;
end;

procedure TfrmMain.ApplicationEvents1Activate(Sender: TObject);
begin
  hide;
end;

procedure TfrmMain.WMSyscommand(var message: TWMMouse);
begin
  inherited;
  if message.Keys=SC_MINIMIZE then hide;
  message.Result:=-1;
end;

procedure TfrmMain.ToolButton7Click(Sender: TObject);
begin
  if MakeDBConn then ConnectString:=GetConnectString;
end;

procedure TfrmMain.UpdateConfig;
var
  INI:tinifile;
  autorun:boolean;
begin
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));

  autorun:=ini.readBool(IniSection,'开机自动运行',false);
  ifRecLog:=ini.readBool(IniSection,'调试日志',false);

  GroupName:=trim(ini.ReadString(IniSection,'工作组',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'仪器字母','')));//读出来是大写就万无一失了
  SpecType:=ini.ReadString(IniSection,'默认样本类型','');
  SpecStatus:=ini.ReadString(IniSection,'默认样本状态','');
  CombinID:=ini.ReadString(IniSection,'组合项目代码','');

  LisFormCaption:=ini.ReadString(IniSection,'检验系统窗体标题','');

  QuaContSpecNoG:=ini.ReadString(IniSection,'高值质控联机号','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'常值质控联机号','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'低值质控联机号','9997');

  MrConnStr:=ini.ReadString(IniSection,'连接仪器数据库','');

  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);

  try
    ADOConn_BS.Connected := false;
    ADOConn_BS.ConnectionString := MrConnStr;
    ADOConn_BS.Connected := true;
    ifConnSucc:=true;
  except
    ifConnSucc:=false;
    showmessage('连接仪器数据库失败!');
  end;
end;

function TfrmMain.LoadInputPassDll: boolean;
TYPE
    TDLLFUNC=FUNCTION:boolean;
VAR
    HLIB:THANDLE;
    DLLFUNC:TDLLFUNC;
    PassFlag:boolean;
begin
    result:=false;
    HLIB:=LOADLIBRARY('OnOffLogin.dll');
    IF HLIB=0 THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    DLLFUNC:=TDLLFUNC(GETPROCADDRESS(HLIB,'showfrmonofflogin'));
    IF @DLLFUNC=NIL THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    PassFlag:=DLLFUNC;
    FREELIBRARY(HLIB);
    result:=passflag;
end;

function TfrmMain.MakeDBConn:boolean;
var
  newconnstr,ss: string;
  Label labReadIni;
begin
  result:=false;

  labReadIni:
  newconnstr := GetConnectString;
  
  try
    ADOConnection1.Connected := false;
    ADOConnection1.ConnectionString := newconnstr;
    ADOConnection1.Connected := true;
    result:=true;
  except
  end;
  if not result then
  begin
    ss:='服务器'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '数据库'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '集成登录模式'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '用户'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '口令'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('连接数据库','连接数据库',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
begin
  if LoadInputPassDll then
  begin
    ss:='连接仪器数据库'+#2+'DBConn'+#2+#2+'1'+#2+#2+#3+
      '工作组'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '仪器字母'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '检验系统窗体标题'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本类型'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本状态'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '组合项目代码'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '开机自动运行'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '调试日志'+#2+'CheckListBox'+#2+#2+'0'+#2+'注:强烈建议在正常运行时关闭'+#2+#3+
      '高值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '常值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '低值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
  end;
end;

procedure TfrmMain.ToolButton5Click(Sender: TObject);
var
  ss:string;
begin
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'将该窗体标题栏上的字符串发给开发者,以获取注册码'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('注册:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
    if ifRegister then bRegister:=true else bRegister:=false;
end;

procedure TfrmMain.BitBtn3Click(Sender: TObject);
VAR
  FInts:OleVariant;
  
  adotemp22:tadoquery;
  ReceiveItemInfo:OleVariant;
begin
  if not ifConnSucc then
  begin
    showmessage('连接仪器数据库失败,不能发送!');
    exit;
  end;
  
  (sender as TBitBtn).Enabled:=false;  

  adotemp22:=tadoquery.Create(nil);
  adotemp22.Connection:=ADOConn_BS;
  adotemp22.Close;
  adotemp22.SQL.Clear;
  adotemp22.SQL.Text:='select orderid,sampleid,patientid,acqutime,picturepath,'+
                      'WBC,NIT,URO,PRO,pH,BLD,SG,BIL,Vc,KET,GLU,Color,Turbidity,MCa,Ca,CRE,redCell,whiteCell,whiteCellgroup,squa,'+
                      'nonSqua,otherSqua,cylinder,hyalineCast,granularCast,crystal,urateCrystal,otherCrystal,speram,baterium,yeast,'+
                      'mucus,fatBall,trichmo,resultchange,unredcell '+
                      ' from UrineResult '+
                      ' where acqudate='''+FormatDateTime('YYYYMMDD',DateTimePicker1.Date)+''' ';
  adotemp22.Open;
  while not adotemp22.Eof do
  begin
    orderid:=adotemp22.fieldbyname('orderid').AsString;
    sampleid:=adotemp22.fieldbyname('sampleid').AsString;
    patientid:=adotemp22.fieldbyname('patientid').AsString;
    acqutime:=adotemp22.fieldbyname('acqutime').AsString;
    picturepath:=adotemp22.fieldbyname('picturepath').AsString;
  
    ReceiveItemInfo:=VarArrayCreate([0,36+19-1],varVariant);//19个图片项目

    ReceiveItemInfo[0]:=VarArrayof(['WBC',StringReplace(copy(adotemp22.fieldbyname('WBC').AsString,5,MaxInt),'Cells/uL','',[rfReplaceAll, rfIgnoreCase]),'','']);
    ReceiveItemInfo[1]:=VarArrayof(['NIT',copy(adotemp22.fieldbyname('NIT').AsString,5,MaxInt),'','']);
    ReceiveItemInfo[2]:=VarArrayof(['URO',StringReplace(copy(adotemp22.fieldbyname('URO').AsString,5,MaxInt),'umol/L','',[rfReplaceAll, rfIgnoreCase]),'','']);
    ReceiveItemInfo[3]:=VarArrayof(['PRO',StringReplace(copy(adotemp22.fieldbyname('PRO').AsString,5,MaxInt),'g/L','',[rfReplaceAll, rfIgnoreCase]),'','']);
    ReceiveItemInfo[4]:=VarArrayof(['pH',copy(adotemp22.fieldbyname('pH').AsString,5,MaxInt),'','']);
    ReceiveItemInfo[5]:=VarArrayof(['BLD',StringReplace(copy(adotemp22.fieldbyname('BLD').AsString,5,MaxInt),'Cells/uL','',[rfReplaceAll, rfIgnoreCase]),'','']);
    ReceiveItemInfo[6]:=VarArrayof(['SG',copy(adotemp22.fieldbyname('SG').AsString,5,MaxInt),'','']);
    ReceiveItemInfo[7]:=VarArrayof(['BIL',StringReplace(copy(adotemp22.fieldbyname('BIL').AsString,5,MaxInt),'umol/L','',[rfReplaceAll, rfIgnoreCase]),'','']);
    ReceiveItemInfo[8]:=VarArrayof(['Vc',StringReplace(copy(adotemp22.fieldbyname('Vc').AsString,5,MaxInt),'mmol/L','',[rfReplaceAll, rfIgnoreCase]),'','']);
    ReceiveItemInfo[9]:=VarArrayof(['KET',StringReplace(copy(adotemp22.fieldbyname('KET').AsString,5,MaxInt),'mmol/L','',[rfReplaceAll, rfIgnoreCase]),'','']);
    ReceiveItemInfo[10]:=VarArrayof(['GLU',StringReplace(copy(adotemp22.fieldbyname('GLU').AsString,5,MaxInt),'mmol/L','',[rfReplaceAll, rfIgnoreCase]),'','']);
    ReceiveItemInfo[11]:=VarArrayof(['Color',adotemp22.fieldbyname('Color').AsString,'','']);
    ReceiveItemInfo[12]:=VarArrayof(['Turbidity',adotemp22.fieldbyname('Turbidity').AsString,'','']);
    ReceiveItemInfo[13]:=VarArrayof(['MCa',adotemp22.fieldbyname('MCa').AsString,'','']);
    ReceiveItemInfo[14]:=VarArrayof(['Ca',adotemp22.fieldbyname('Ca').AsString,'','']);
    ReceiveItemInfo[15]:=VarArrayof(['CRE',adotemp22.fieldbyname('CRE').AsString,'','']);
    ReceiveItemInfo[16]:=VarArrayof(['redCell',adotemp22.fieldbyname('redCell').AsString,'','']);
    ReceiveItemInfo[17]:=VarArrayof(['whiteCell',adotemp22.fieldbyname('whiteCell').AsString,'','']);
    ReceiveItemInfo[18]:=VarArrayof(['whiteCellgroup',adotemp22.fieldbyname('whiteCellgroup').AsString,'','']);
    ReceiveItemInfo[19]:=VarArrayof(['squa',adotemp22.fieldbyname('squa').AsString,'','']);
    ReceiveItemInfo[20]:=VarArrayof(['nonSqua',adotemp22.fieldbyname('nonSqua').AsString,'','']);
    ReceiveItemInfo[21]:=VarArrayof(['otherSqua',adotemp22.fieldbyname('otherSqua').AsString,'','']);
    ReceiveItemInfo[22]:=VarArrayof(['cylinder',adotemp22.fieldbyname('cylinder').AsString,'','']);
    ReceiveItemInfo[23]:=VarArrayof(['hyalineCast',adotemp22.fieldbyname('hyalineCast').AsString,'','']);
    ReceiveItemInfo[24]:=VarArrayof(['granularCast',adotemp22.fieldbyname('granularCast').AsString,'','']);
    ReceiveItemInfo[25]:=VarArrayof(['crystal',adotemp22.fieldbyname('crystal').AsString,'','']);
    ReceiveItemInfo[26]:=VarArrayof(['urateCrystal',adotemp22.fieldbyname('urateCrystal').AsString,'','']);
    ReceiveItemInfo[27]:=VarArrayof(['otherCrystal',adotemp22.fieldbyname('otherCrystal').AsString,'','']);
    ReceiveItemInfo[28]:=VarArrayof(['speram',adotemp22.fieldbyname('speram').AsString,'','']);
    ReceiveItemInfo[29]:=VarArrayof(['baterium',adotemp22.fieldbyname('baterium').AsString,'','']);
    ReceiveItemInfo[30]:=VarArrayof(['yeast',adotemp22.fieldbyname('yeast').AsString,'','']);
    ReceiveItemInfo[31]:=VarArrayof(['mucus',adotemp22.fieldbyname('mucus').AsString,'','']);
    ReceiveItemInfo[32]:=VarArrayof(['fatBall',adotemp22.fieldbyname('fatBall').AsString,'','']);
    ReceiveItemInfo[33]:=VarArrayof(['trichmo',adotemp22.fieldbyname('trichmo').AsString,'','']);
    ReceiveItemInfo[34]:=VarArrayof(['resultchange',adotemp22.fieldbyname('resultchange').AsString,'','']);
    ReceiveItemInfo[35]:=VarArrayof(['unredcell',adotemp22.fieldbyname('unredcell').AsString,'','']);

    ReceiveItemInfo[36]:=VarArrayof(['100X-01','','',picturepath+'\100X\100X-01.jpg']);
    ReceiveItemInfo[37]:=VarArrayof(['100X-02','','',picturepath+'\100X\100X-02.jpg']);
    ReceiveItemInfo[38]:=VarArrayof(['100X-03','','',picturepath+'\100X\100X-03.jpg']);
    ReceiveItemInfo[39]:=VarArrayof(['100X-04','','',picturepath+'\100X\100X-04.jpg']);
    ReceiveItemInfo[40]:=VarArrayof(['100X-05','','',picturepath+'\100X\100X-05.jpg']);
    ReceiveItemInfo[41]:=VarArrayof(['100X-06','','',picturepath+'\100X\100X-06.jpg']);
    ReceiveItemInfo[42]:=VarArrayof(['100X-07','','',picturepath+'\100X\100X-07.jpg']);
    ReceiveItemInfo[43]:=VarArrayof(['100X-08','','',picturepath+'\100X\100X-08.jpg']);
    ReceiveItemInfo[44]:=VarArrayof(['100X-09','','',picturepath+'\100X\100X-09.jpg']);
    ReceiveItemInfo[45]:=VarArrayof(['400X-01','','',picturepath+'\400X\400X-01.jpg']);
    ReceiveItemInfo[46]:=VarArrayof(['400X-02','','',picturepath+'\400X\400X-02.jpg']);
    ReceiveItemInfo[47]:=VarArrayof(['400X-03','','',picturepath+'\400X\400X-03.jpg']);
    ReceiveItemInfo[48]:=VarArrayof(['400X-04','','',picturepath+'\400X\400X-04.jpg']);
    ReceiveItemInfo[49]:=VarArrayof(['400X-05','','',picturepath+'\400X\400X-05.jpg']);
    ReceiveItemInfo[50]:=VarArrayof(['400X-06','','',picturepath+'\400X\400X-06.jpg']);
    ReceiveItemInfo[51]:=VarArrayof(['400X-07','','',picturepath+'\400X\400X-07.jpg']);
    ReceiveItemInfo[52]:=VarArrayof(['400X-08','','',picturepath+'\400X\400X-08.jpg']);
    ReceiveItemInfo[53]:=VarArrayof(['400X-09','','',picturepath+'\400X\400X-09.jpg']);
    ReceiveItemInfo[54]:=VarArrayof(['RedCellPhase','','',picturepath+'\RedCellPhase.jpg']);

    if bRegister then
    begin
      FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
      FInts.fData2Lis(ReceiveItemInfo,rightstr('0000'+orderid,4),
        FormatDateTime('YYYY-MM-DD',DateTimePicker1.Date)+' '+copy(acqutime,1,2)+':'+copy(acqutime,3,2)+':'+copy(acqutime,5,2),
        (GroupName),(SpecType),(SpecStatus),(EquipChar),
        (CombinID),'',
        (LisFormCaption),(ConnectString),
        (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
        ifRecLog,true,'常规');
      if not VarIsEmpty(FInts) then FInts:= unAssigned;
    end;

    adotemp22.Next;
  end;

  adotemp22.Free;
  
  (sender as TBitBtn).Enabled:=true;
end;

initialization
    hnd := CreateMutex(nil, True, Pchar(ExtractFileName(Application.ExeName)));
    if GetLastError = ERROR_ALREADY_EXISTS then
    begin
        MessageBox(application.Handle,pchar('该程序已在运行中！'),
                    '系统提示',MB_OK+MB_ICONinformation);
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.

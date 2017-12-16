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

uses ucommfunction, USearchFile;

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

  hnd:integer;
  bRegister:boolean;

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

procedure AFindCallBack(const filename:string;const info:tsearchrec;var quit:boolean);
var
  ls,lsValue,sList:tstrings;
  i:integer;

  SpecNo:string;
  FInts:OleVariant;
  ReceiveItemInfo:OleVariant;

  ini:Tinifile;

  //图形路径
  HPLT:string;
  HRBC:string;
  HWBC:string;
  SBASO:string;
  SDIFF:string;
  SIMI:string;
  SNRBC:string;
  SPLT:string;
  SRET:string;
  SRET_E:string;

  //HRBCY:string;
  //HWDFY:string;
  //SPLT_F:string;
  //SPLT_O:string;
  //SWDF:string;
  //SWNR:string;
  //SWPC:string;
  //=========

  s1:string;
  i0:TDateTime;//上次检验时间
  i1:TDateTime;//本次检验时间
  sName:string;//文件名
  fs:TFormatSettings;
  s2:string;
begin
  {sName:=ExtractFileName(filename);
  
  sList:=TStringList.Create;
  ExtractStrings(['_'],[],PChar(sName),sList);
  if sList.Count<2 then begin sList.Free;exit;end;
  s1:=sList[0]+'_'+sList[1];
  sList.Free;
    
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  i0:=ini.ReadDateTime(FormatDateTime('YYYYMMDD',now),s1,0);
  ini.Free;

  ls:=Tstringlist.Create;
  ls.LoadFromFile(filename);
  if ls.Count<=0 then begin ls.Free;exit;end;//如果仪器还没向cdf文件中写完，则等待写完

  //本次检验时间
  i1:=1;
  for i :=0  to ls.Count-1 do
  begin
    lsValue:=StrToList(ls[i],big_result);//将每行导入到字符串列表中

    if lsValue.Count<20 then begin lsValue.Free;continue;end;
    s2:=StringReplace(lsValue[19],'/','-',[rfReplaceAll, rfIgnoreCase]);

    if lsValue[0]<>'00' then begin lsValue.Free;continue;end;

    fs.DateSeparator:='-';
    fs.TimeSeparator:=':';
    fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
    i1:=StrtoDateTimeDef(s2,i1,fs);

    lsValue.Free;
  end;
  //==========

  if i1<=i0 then begin ls.Free;exit;end;//该文件已经处理过或已处理过以前做的
  
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  ini.WriteDateTime(FormatDateTime('YYYYMMDD',now),s1,i1);
  ini.Free;

  if length(frmMain.memo1.Lines.Text)>=60000 then frmMain.memo1.Lines.Clear;//memo只能接受64K个字符
  frmMain.memo1.Lines.Add(filename);

  //取图形数据
  for i :=0  to ls.Count-1 do
  begin
    lsValue:=StrToList(ls[i],big_result);//将每行导入到字符串列表中

    if lsValue.Count<4 then continue;

    if uppercase(lsValue[2])='HPLT' then HPLT:=lsValue[3];
    if uppercase(lsValue[2])='HRBC' then HRBC:=lsValue[3];
    if uppercase(lsValue[2])='HWBC' then HWBC:=lsValue[3];
    if uppercase(lsValue[2])='SBASO' then SBASO:=lsValue[3];
    if uppercase(lsValue[2])='SDIFF' then SDIFF:=lsValue[3];
    if uppercase(lsValue[2])='SIMI' then SIMI:=lsValue[3];
    if uppercase(lsValue[2])='SNRBC' then SNRBC:=lsValue[3];
    if uppercase(lsValue[2])='SPLT' then SPLT:=lsValue[3];
    if uppercase(lsValue[2])='SRET' then SRET:=lsValue[3];
    if uppercase(lsValue[2])='SRET-E' then SRET_E:=lsValue[3];

    //if uppercase(lsValue[2])='HRBCY' then HRBCY:=lsValue[3];
    //if uppercase(lsValue[2])='HWDFY' then HWDFY:=lsValue[3];
    //if uppercase(lsValue[2])='SPLT-F' then SPLT_F:=lsValue[3];
    //if uppercase(lsValue[2])='SPLT-O' then SPLT_O:=lsValue[3];
    //if uppercase(lsValue[2])='SWDF' then SWDF:=lsValue[3];
    //if uppercase(lsValue[2])='SWNR' then SWNR:=lsValue[3];
    //if uppercase(lsValue[2])='SWPC' then SWPC:=lsValue[3];

    lsValue.Free;
  end;
  //============

  ReceiveItemInfo:=VarArrayCreate([0,ls.Count-1],varVariant);

  for i :=0  to ls.Count-1 do
  begin
    lsValue:=StrToList(ls[i],big_result);//将每行导入到字符串列表中

    if lsValue.Count<4 then
    begin
      ReceiveItemInfo[i]:=VarArrayof(['','','','']);
      continue;
    end;

    if lsValue[0]='0' then SpecNo:=rightstr('0000'+lsValue[3],4);

    if lsValue[0]='1' then
    begin
      if uppercase(lsValue[1])='PLT' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',HPLT])
      else if uppercase(lsValue[1])='RBC' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',HRBC])
      else if uppercase(lsValue[1])='WBC' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',HWBC])
      else if uppercase(lsValue[1])='BASO#' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SBASO])
      else if uppercase(lsValue[1])='MPV' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SDIFF])
      else if uppercase(lsValue[1])='MONO#' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SIMI])
      else if uppercase(lsValue[1])='NRBC#' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SNRBC])
      else if uppercase(lsValue[1])='HCT' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SPLT])
      else if uppercase(lsValue[1])='RET#' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SRET])
      else if uppercase(lsValue[1])='RET%' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SRET_E])
      else ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'','']);
    end
    else if lsValue[0]='3' then ReceiveItemInfo[i]:=VarArrayof([lsValue[2],'','',lsValue[3]])
    else ReceiveItemInfo[i]:=VarArrayof(['','','','']);

    lsValue.Free;
  end;
  
  ls.Free;

  if bRegister then
  begin
    FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
    FInts.fData2Lis(ReceiveItemInfo,(SpecNo),'',
      (GroupName),(SpecType),(SpecStatus),(EquipChar),
      (CombinID),'',(LisFormCaption),(ConnectString),
      (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
      ifRecLog,true,'常规');
    if not VarIsEmpty(FInts) then FInts:= unAssigned;
  end;//}
end;

procedure TfrmMain.BitBtn3Click(Sender: TObject);
VAR
  adotemp22,adotemp,adotemp33:tadoquery;
  SamNo:string;
  ReceiveItemInfo:OleVariant;
  FInts:OleVariant;
  sName,sSex,sAge,sKB,sBQ,sBLH,sBedNo,sLCZD,sSJYS,sJYYS:String;
  i,RecNum:integer;

  picturepath:string;
  qqq:boolean;
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
    picturepath:=adotemp22.fieldbyname('picturepath').AsString;

    qqq:=false;
    findfile(qqq,picturepath,'*.jpg',AFindCallBack,true,true);

    
    adotemp33:=tadoquery.Create(nil);
    adotemp33.Connection:=ADOConn_BS;
    adotemp33.Close;
    adotemp33.SQL.Clear;
    adotemp33.SQL.Text:='select count(*) as RecNum from Visc where TestDataID='+adotemp22.fieldbyname('TestDataID').AsString;
    adotemp33.Open;
    RecNum:=adotemp33.fieldbyname('RecNum').AsInteger;
    adotemp33.Free;
  
    ReceiveItemInfo:=VarArrayCreate([0,38+RecNum-1],varVariant);
    
    adotemp:=tadoquery.Create(nil);
    adotemp.Connection:=ADOConn_BS;
    adotemp.Close;
    adotemp.SQL.Clear;
    adotemp.SQL.Text:='select ShearRate,Visc from Visc where TestDataID='+adotemp22.fieldbyname('TestDataID').AsString;
    adotemp.Open;
    i:=0;
    while not adotemp.Eof do
    begin
      ReceiveItemInfo[i]:=VarArrayof([adotemp.fieldbyname('ShearRate').AsString,adotemp.fieldbyname('Visc').AsString,'','']);
      inc(i);
      adotemp.Next;
    end;
    adotemp.Free;

    SamNo:=adotemp22.fieldbyname('序号').AsString;
    sName:=adotemp22.fieldbyname('姓名').AsString;
    sSex:=ifThen(uppercase(adotemp22.fieldbyname('性别').AsString)='TRUE','男','女');
    sAge:=adotemp22.fieldbyname('年龄').AsString;
    sKB:=adotemp22.fieldbyname('科别').AsString;
    sBQ:=adotemp22.fieldbyname('病区').AsString;
    sBLH:=adotemp22.fieldbyname('病历号').AsString;
    sBedNo:=adotemp22.fieldbyname('床号').AsString;
    sLCZD:=adotemp22.fieldbyname('临床诊断').AsString;
    sSJYS:=adotemp22.fieldbyname('送检医生').AsString;
    sJYYS:=adotemp22.fieldbyname('检验医生').AsString;
      
    ReceiveItemInfo[0+i]:=VarArrayof(['全血粘度',adotemp22.fieldbyname('全血粘度').AsString,'','']);
    ReceiveItemInfo[1+i]:=VarArrayof(['血浆粘度',adotemp22.fieldbyname('血浆粘度').AsString,'','']);
    ReceiveItemInfo[2+i]:=VarArrayof(['压积',adotemp22.fieldbyname('压积').AsString,'','']);
    ReceiveItemInfo[3+i]:=VarArrayof(['血沉',adotemp22.fieldbyname('血沉').AsString,'','']);
    ReceiveItemInfo[4+i]:=VarArrayof(['血沉最大沉降率',adotemp22.fieldbyname('血沉最大沉降率').AsString,'','']);
    ReceiveItemInfo[5+i]:=VarArrayof(['血沉最大沉降率时间',adotemp22.fieldbyname('血沉最大沉降率时间').AsString,'','']);
    ReceiveItemInfo[6+i]:=VarArrayof(['全血低切相对指数',adotemp22.fieldbyname('全血低切相对指数').AsString,'','']);
    ReceiveItemInfo[7+i]:=VarArrayof(['全血高切相对指数',adotemp22.fieldbyname('全血高切相对指数').AsString,'','']);
    ReceiveItemInfo[8+i]:=VarArrayof(['血沉方程K值',adotemp22.fieldbyname('血沉方程K值').AsString,'','']);
    ReceiveItemInfo[9+i]:=VarArrayof(['红细胞聚集指数',adotemp22.fieldbyname('红细胞聚集指数').AsString,'','']);
    ReceiveItemInfo[10+i]:=VarArrayof(['红细胞聚集系数',adotemp22.fieldbyname('红细胞聚集系数').AsString,'','']);
    ReceiveItemInfo[11+i]:=VarArrayof(['红细胞变形指数',adotemp22.fieldbyname('红细胞变形指数').AsString,'','']);
    ReceiveItemInfo[12+i]:=VarArrayof(['全血低切还原粘度',adotemp22.fieldbyname('全血低切还原粘度').AsString,'','']);
    ReceiveItemInfo[13+i]:=VarArrayof(['全血高切还原粘度',adotemp22.fieldbyname('全血高切还原粘度').AsString,'','']);
    ReceiveItemInfo[14+i]:=VarArrayof(['红细胞变形指数TK',adotemp22.fieldbyname('红细胞变形指数TK').AsString,'','']);
    ReceiveItemInfo[15+i]:=VarArrayof(['红细胞刚性指数',adotemp22.fieldbyname('红细胞刚性指数').AsString,'','']);
    ReceiveItemInfo[16+i]:=VarArrayof(['卡松粘度',adotemp22.fieldbyname('卡松粘度').AsString,'','']);
    ReceiveItemInfo[17+i]:=VarArrayof(['血红蛋白',adotemp22.fieldbyname('血红蛋白').AsString,'','']);
    ReceiveItemInfo[18+i]:=VarArrayof(['红细胞内粘度',adotemp22.fieldbyname('红细胞内粘度').AsString,'','']);
    ReceiveItemInfo[19+i]:=VarArrayof(['低切流阻',adotemp22.fieldbyname('低切流阻').AsString,'','']);
    ReceiveItemInfo[20+i]:=VarArrayof(['中切流阻',adotemp22.fieldbyname('中切流阻').AsString,'','']);
    ReceiveItemInfo[21+i]:=VarArrayof(['高切流阻',adotemp22.fieldbyname('高切流阻').AsString,'','']);
    ReceiveItemInfo[22+i]:=VarArrayof(['纤维蛋白原',adotemp22.fieldbyname('纤维蛋白原').AsString,'','']);
    ReceiveItemInfo[23+i]:=VarArrayof(['血胆固醇',adotemp22.fieldbyname('血胆固醇').AsString,'','']);
    ReceiveItemInfo[24+i]:=VarArrayof(['甘油三脂',adotemp22.fieldbyname('甘油三脂').AsString,'','']);
    ReceiveItemInfo[25+i]:=VarArrayof(['高密脂蛋白',adotemp22.fieldbyname('高密脂蛋白').AsString,'','']);
    ReceiveItemInfo[26+i]:=VarArrayof(['血糖',adotemp22.fieldbyname('血糖').AsString,'','']);
    ReceiveItemInfo[27+i]:=VarArrayof(['血小板粘附率',adotemp22.fieldbyname('血小板粘附率').AsString,'','']);
    ReceiveItemInfo[28+i]:=VarArrayof(['体外血栓干重',adotemp22.fieldbyname('体外血栓干重').AsString,'','']);
    ReceiveItemInfo[29+i]:=VarArrayof(['红细胞电泳',adotemp22.fieldbyname('红细胞电泳').AsString,'','']);
    ReceiveItemInfo[30+i]:=VarArrayof(['血小板聚集率',adotemp22.fieldbyname('血小板聚集率').AsString,'','']);
    ReceiveItemInfo[31+i]:=VarArrayof(['体外血栓长度',adotemp22.fieldbyname('体外血栓长度').AsString,'','']);
    ReceiveItemInfo[32+i]:=VarArrayof(['结果分析',adotemp22.fieldbyname('结果分析').AsString,'','']);
    ReceiveItemInfo[33+i]:=VarArrayof(['全血中切还原粘度',adotemp22.fieldbyname('全血中切还原粘度').AsString,'','']);
    ReceiveItemInfo[34+i]:=VarArrayof(['屈服应力',adotemp22.fieldbyname('屈服应力').AsString,'','']);
    ReceiveItemInfo[35+i]:=VarArrayof(['红细胞电泳指数',adotemp22.fieldbyname('红细胞电泳指数').AsString,'','']);
    ReceiveItemInfo[36+i]:=VarArrayof(['全血中切相对指数',adotemp22.fieldbyname('全血中切相对指数').AsString,'','']);
    ReceiveItemInfo[37+i]:=VarArrayof(['红细胞计数',adotemp22.fieldbyname('红细胞计数').AsString,'','']);

    if bRegister then
    begin
      FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
      FInts.fData2Lis(ReceiveItemInfo,rightstr('0000'+SamNo,4),
        FormatDateTime('YYYY-MM-DD',DateTimePicker1.Date)+' '+FormatDateTime('hh:nn:ss',adotemp22.fieldbyname('时间').AsDateTime),
        (GroupName),(SpecType),(SpecStatus),(EquipChar),
        (CombinID),
        sName+'{!@#}'+sSex+'{!@#}{!@#}'+sAge+'{!@#}'+sBLH+'{!@#}'+sKB+'{!@#}'+sSJYS+'{!@#}'+sBedNo+'{!@#}'+sLCZD+'{!@#}{!@#}'+sJYYS,
        (LisFormCaption),(ConnectString),
        (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
        true,true,'常规');
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

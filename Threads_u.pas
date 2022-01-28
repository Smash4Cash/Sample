unit Threads_u;

interface

uses
  Windows, Classes, SysUtils, Graphics, StdCtrls, IdFTP, IdException, IdExceptionCore,
  DevProp_u, Chan_u, Dev_u, dllimport_u;

type
  TData1 = record
    DT: TDateTime;
    B: byte;
    C: integer;
    V: double;
  end;
  TData2 = record
    DT: TDateTime;
    B: byte;
    C: integer;
    Integr,Max,Aver,Min: double;
  end;

  // ������ ������ ��������� �� COM-�����
  TCOMThread = class(TThread)
  Index       : integer;
  TreeMutex1,
  COMh        : THandle;
  Protocol    : TProtocolType;
  private
    ChanToFill: TChan;
  protected
    procedure FillMeasure;
    procedure Execute; override;
  end;

  // ����� ��������������� TCP �������
  TArch = class(TThread)
  TreeMutex2  : THandle;
  private
    DevToFill: TDev;
  protected
    procedure SetControls;
    procedure FillArch;
    procedure Execute; override;
    procedure SetBtnCopyEnabled;
    procedure SetBtnCopyDisabled;
  end;

  // ����� ������� ����������� TCP �������
  TArchManual = class(TThread)
  TreeMutex2  : THandle;
  private
  public
    Dev       : TDev;
    Rewrite   : TRewrite;
  protected
    procedure Execute; override;
    procedure FillArch;
  end;

  // ����� �������� �� ����� RS-�������
  TArcLoad = class(TThread)
  private
    arD: array of TMeasData;
  public
    Chan        : TChan;
    strToday    : shortstring;
    fTerminated : boolean;
  protected
    procedure Fill;
    procedure Execute; override;
    procedure Clear;
    procedure Finish;
    procedure Viz;
    procedure UnViz;
  end;

  // ����� �������� �� TCP ������ �������
  TTreeLoadTCP = class(TThread)
  private
  public
    Dev         : TDev;
    fTerminated : boolean;
    Err         : integer;
  protected
    procedure Fill;
    procedure Execute; override;
    procedure Start;
    procedure Finish;
  end;

  // ����� ����������� ����� �� FTP
  TFTPGetThread = class(TThread)
  private
    FStream: TFileStream;
  public
    IdFTP: TIdFTP;
    SourcePath, TargetPath: String;
    bResume, bDone, bNew: boolean;
    Mess: string;
    procedure StreamFree;
  protected
    procedure Execute; override;
  end;

var
  thArcLoad     : TArcLoad;
  thTreeLoadTCP : TTreeLoadTCP;
  ThGet         : TFTPGetThread;

const
  TCP_ERROR_BAD_NET_PATH = 53;

implementation

uses ElTree, Main_u, IEEE754, Utils_u, Series, COM1_u, DirTree_u, Ping_u;

procedure TCOMThread.FillMeasure;
begin
  if Assigned(ChanToFill) then
    ChanToFill.FillMeasureCurrent;
end;

// ����� ������ ��������� �� COM-�����
procedure TCOMThread.Execute;
var
  Err            : longint;
  i,j,ind, Code  : integer;
  sV             : shortstring;
  Item, ChanItem : TElTreeItem;
  V              : double;
  DT, DTnow      : TDateTime;
  Dev, DevM      : TDev;
  Chan, ChanM    : TChan;
  dtMin, dtI     : TDateTime;
  function Compare(_Item: TElTreeItem; _Chan: TChan): boolean;
  begin
    Result := false;
    if Assigned(_Item) then
      if (_Item.Level=2) then
        if Assigned(_Item.Data) then
          if TChan(_Item.Data) = ChanM then
            Result := true;
  end;
begin
  FreeOnTerminate:=false; ChanToFill := nil;
  repeat
    if (Index >= 0)and(Index < lstCOM.Count) then
    begin
      if WaitForSingleObject(TreeMutex1,20) = WAIT_OBJECT_0 then
      begin
        dtMin := High(INT64); ChanM := nil; ChanItem := nil;
        for i := 0 to lstDevice.Count-1 do
        begin
          // ����� ������, ������� ������� ������ � 1-�� �������
          if TDev(lstDevice[i]).GetCOMIndex = Index then
          begin
            Dev := TDev(lstDevice[i]);
            for j:=0 to Length(Dev.arChan)-1 do
            begin
              Chan := TChan(Dev.arChan[j]);
              if Chan.Checked <> cbUnchecked then
              begin
                dtI := Chan.lastDT + Chan.Period/86400;
                if (dtI < dtMin) then
                begin
                  DevM  := Dev;
                  ChanM := Chan;
                  dtMin := dtI;
                end;
              end;
            end;
          end;
        end;
        DTnow := Now;
        if (Assigned(DevM))and(Assigned(ChanM)) then if ((ChanM.lastDT + ChanM.Period/86400) <= DTnow)or(ChanM.lastDT=0) then
        begin
          case Protocol of
            ptASCII: Err := DevM.MeasureRead(COMh,ChanM.Address,V);
            ptModBus: Err := DevM.MeasureReadMB(COMh,ChanM.Address,V);
          end;
//          ChanM.lastDT := Now;
          ind := ChanM.GetTreeItemIndex;
          if (ind >=0)and(ind < frmMain.NetTree.Items.Count) then
            ChanItem := frmMain.NetTree.Items[ind];
          if ChanM.Checked <> cbUnchecked then
          begin
            if Err = EX_OK then
            begin
              sV := RealToStr(V, SymbolsAfterPointsForMeasureValue);
              DT := Now;
              if V <> ErrValue then
              begin
                ChanM.AddMeasureData(DT, V, Err);
              end
              else begin
                ChanM.lastErr := EX_DATA_ERROR;
                ChanM.lastDT := DT;
                sV := '#'+IntToStr(Abs(EX_DATA_ERROR));;
              end;
              // ���� ������ ����� ������ ����, �� ������� ����� �� �����
              if Compare(frmMain.NetTree.Selected, ChanM) then
              begin
                ChanToFill := ChanM;
                Synchronize(FillMeasure);
              end;
            end
            else begin
              ChanM.lastErr := Err;
              ChanM.lastDT := DT;
              // ���� ������ ����� ������ ����, �� ������� ����� �� �����
              if Compare(frmMain.NetTree.Selected, ChanM) then
              begin
                ChanToFill := ChanM;
                Synchronize(FillMeasure);
              end;
              if Err > 0 then
                sV := '$'+IntToStr(Abs(Err))
              else
                sV := '#'+IntToStr(Abs(Err));
            end;
            ChanItem.ColumnText.Strings[0] := sV;
          end;
        end;
        ReleaseMutex(TreeMutex1);
        Sleep(30);
      end;
    end;
    Sleep(30);
  until (Terminated);
end;

procedure TArch.SetControls;
begin
  if Assigned(DevToFill) then
    with DevToFill do with frmArchiveTCP do
    begin
      btnStop.Visible := (esArchive in ExchStatus);
      btnCopy.Visible := not(esArchive in ExchStatus);
//      btnCopy.Enabled := not(esAutoCopy in ExchStatus);
    end;
end;

procedure TArch.SetBtnCopyEnabled;
begin
  frmArchiveTCP.btnCopy.Enabled := true;
end;

procedure TArch.SetBtnCopyDisabled;
begin
  frmArchiveTCP.btnCopy.Enabled := false;
end;

procedure TArch.FillArch;
begin
  if Assigned(DevToFill) then
    with DevToFill do
    begin
      TCPFillArch;
      TCPFillArchContrls;
    end;
end;

// ����� ��������������� TCP �������
procedure TArch.Execute;
var
  Err       : longint;
  i,j,ind, Code : integer;
  sV        : shortstring;
  Item, DevItem : TElTreeItem;
  V         : double;
  DT        : TDateTime;
  Dev       : TDev;
  DW        : dword;
  R         : boolean;
begin
  FreeOnTerminate:=false;
  repeat
    if frmMain.AutoCopy then
    begin
      if WaitForSingleObject(TreeMutex2,20) = WAIT_OBJECT_0 then
      begin
        for i := 0 to lstDeviceTCP.Count-1 do
        begin
          Dev := TDev(lstDeviceTCP[i]);
          for j:=0 to frmMain.TCPTree.Items.Count-1 do
          begin
            Item := frmMain.TCPTree.Items[j];
            if (Item.Level = 0)and(Item.Checked) then
              if Item.Data = Dev then
                break;
          end;
          ind := j;
          if ind < frmMain.TCPTree.Items.Count then
          begin
            if ((Dev.ArchDT + (frmMain.CopyPeriod/24)) <= Now)or
               ((Dev.ArchErrDT > 0)and((Dev.ArchErrDT + (frmMain.CopyError/1440)) <= Now)) then
            begin
              Include(Dev.ExchStatus, esAutoCopy);
              if Dev = frmMain.GetSelectedDev(frmMain.TCPTree, 0) then
              begin
                DevToFill := Dev;
                Synchronize(SetControls);
              end;
              Synchronize(SetBtnCopyDisabled);
              with frmMain do
                R := Dev.TCPFileRead(Rewrite, GetLastDays, CopyNoToday, CopyReports);
              Exclude(Dev.ExchStatus, esAutoCopy);
              if Dev = frmMain.GetSelectedDev(frmMain.TCPTree, 0) then
              begin
                DevToFill := Dev;
                Synchronize(FillArch);
                Synchronize(SetControls);
              end;
              Synchronize(SetBtnCopyEnabled);
            end;
          end;
        end;
        ReleaseMutex(TreeMutex2);
        Sleep(30);
      end;
    end;
    Sleep(30);
  until (Terminated);
end;

// ����� ������� ����������� TCP �������
procedure TArchManual.Execute;
begin
  if Assigned(Dev) then with Dev do
  begin
    Include(ExchStatus, esArchive);
    with frmMain do
      TCPFileRead(Rewrite, GetLastDays, CopyNoToday, CopyReports);
    Exclude(ExchStatus, esArchive);
    if Dev = frmMain.GetSelectedDev(frmMain.TCPTree,0) then
      Synchronize(FillArch);
  end;
end;

procedure TArchManual.FillArch;
begin
  Dev.TCPFillArch;
  Dev.TCPFillArchContrls;
end;

procedure TArcLoad.Fill;
begin
  Chan.FillTableAndChart;
end;

// ����� �������� �� ����� RS-�������
procedure TArcLoad.Execute;
var
  sZN, sDIR : shortstring;
  searchResult : TSearchRec;
  s : string;
  arData  : array [0..9] of TMeasData; // ����� ��� ������ ������
  L : integer;
  W1,W2,W3,W4,W5,W6,KSw : WORD;
  KSb : byte;
  i, j : integer;
  Date,Time : double;
  DateTime  : TDateTime;
  sDate,sTime : shortstring;
  Count, CountAll : longint;
  DataFile : File of byte;
  DataLength : byte;
  DD : REAL48;
  DD3 : double;
  arV : array [0..3] of byte;
  sV : shortstring;
  ind : longint;
  FType : byte;
  sDirName, sFN: shortstring;
begin
  Synchronize(UnViz);
  SetLength(arD,0); CountAll := 0;
  sDirName := strToday;
  if Assigned(Chan) then with Chan do
  begin
    sDIR := GetArcDIR + sDirName + '\';
    sFN := Name + '.arc';
    Synchronize(Clear);
    AssignFile(DataFile, sDIR + sFN);
    Reset(DataFile);
    DataLength := SizeOf(arData);
    BlockRead(DataFile, arData, SizeOf(arData), L);
    while (L >=DataLength)and(not(Terminated)) do
    begin
      Count := Trunc(SizeOf(arData)/SizeOf(arData[0]));
      SetLength(arD,CountAll+Count);
      for j:=0 to Count-1 do
        arD[CountAll+j] := arData[j];
      CountAll := CountAll + Count;
      BlockRead(DataFile, arData, Length(arData), L);
    end;
    CloseFile(DataFile);
  end;
  if not(Terminated) then Synchronize(Finish);
  if not(Terminated) then Synchronize(Fill);
  SetLength(arD,0);
  Synchronize(Viz);
end;

procedure TArcLoad.Clear;
begin
  with frmArchive.asgData do
  begin
    BeginUpdate;
    ClearRows(1,RowCount-1);
    RowCount := 2;
    EndUpdate;
  end;
  Chan.serArch.Clear;
end;

procedure TArcLoad.UnViz;
begin
  with frmArchive do
  begin
    pnlVizBtn.Visible := false;
    Chart.Visible := false;
    btnExport.Visible := false;
    btnArcPrint.Visible := false;
    asgData.Visible := false;
    lblLoadD.Caption := '�������� ������...';
    lblLoadG.Caption := '�������� ������...';
    Refresh;
  end;
end;

procedure TArcLoad.Viz;
begin
  with frmArchive do
  begin
    pnlVizBtn.Visible := true;
    Chart.Visible := true;
    btnExport.Visible := true;
    btnArcPrint.Visible := true;
    lblLoadD.Caption := '';
    lblLoadG.Caption := '';
    Refresh;
  end;
end;

procedure TArcLoad.Finish;
var
  i,j,LocInd : longint;
  Pr, P      : integer;
  S          : shortstring;
begin
  j := 1;
  P := 0; Pr := P;
  frmArchive.asgData.BeginUpdate;
  // ������� �������
  frmArchive.asgData.ClearRows(1,frmArchive.asgData.RowCount-1);
  frmArchive.asgData.RowCount := 2;
  Chan.serArch.BeginUpdate;
  Chan.serArch.Clear;
  for i := 0 to (Length(arD)-1) do
  begin
    LocInd := Chan.serArch.XValues.Locate(arD[i].DT);
    // ������ � ����� �������� ��� �� ���� (��� ���� ����� ������������)
    if (LocInd < 0)or(arD[i].DT=0) then
    begin
      if (j > 1) then frmArchive.asgData.AddRow;
      if arD[i].DT <> 0 then
        frmArchive.asgData.AllCells[0,j] := DateToStr(arD[i].DT)+' '+TimeToStr(arD[i].DT) // �����
      else
        frmArchive.asgData.AllCells[0,j] := '???';                                        // ???
      frmArchive.asgData.AllCells[1,j] := RealToStr(arD[i].Value,4);   // ���������� ��������
      if (arD[i].Err <> 0) then
      // ���� ������
      begin
        frmArchive.asgData.AllCells[2,j] := '$'+IntToStr(arD[i].Err);  // ������� ��� ������
        if arD[i].DT <> 0 then
          Chan.serArch.AddXY(arD[i].DT,arD[i].Value, '', ErrColor);    // ��������� �� ������ (����.����)
      end
      // ��� ������
      else begin
        frmArchive.asgData.AllCells[2,j] := '-';                       // ������� '-' � ��� ������
        if arD[i].DT <> 0 then
          Chan.serArch.AddXY(arD[i].DT,arD[i].Value);                  // ��������� �� ������
      end;
      inc(j);
    end;
    Pr := 10*Round(10*i/(Length(arD)));
    if (frmMain.GetSelectedChan = Chan) then
    begin
      S := '�������� ������... ' + IntToStr(Pr) + '%';
      frmArchive.lblLoadD.Caption := S;
      frmArchive.lblLoadG.Caption := S;
      case frmArchive.pcArchive.ActivePageIndex of
        0: frmArchive.lblLoadD.Refresh;
        1: frmArchive.lblLoadG.Refresh;
      end;
    end;
  end;
  // + Delta ��� �� ���� ������ �������
  i := 0;
  while 2*i < Chan.serArch.Count do
  begin
    Chan.serArch.AddXY(Chan.serArch.XValue[2*i]+(DeltaFast/(24*60*60)),
                                             Chan.serArch.YValue[2*i],
                                             '',
                                             Chan.serArch.ValueColor[2*i]);
    inc(i);
  end;
  Chan.serArch.EndUpdate;
  frmArchive.asgData.Resize;
  frmArchive.asgData.EndUpdate;
end;

// ����� �������� �� TCP ������ �������
procedure TTreeLoadTCP.Execute;
var
  _Err: integer;
begin
  fBreak := false;
  if Assigned(Dev) then with Dev do
  begin
    Synchronize(Start);
    if Ping(IPToString(IP), nil, frmMain.PingTimeout, frmMain.PingResend) <> 0 then
    begin
      _Err := AccessOpen;
      if _Err = NO_ERROR then _Err := GetArcDirList;
      if _Err = NO_ERROR then _Err := AccessClose else AccessClose;
    end
    else
      _Err := TCP_ERROR_BAD_NET_PATH; // 53 = ERROR_BAD_NET_PATH (�� ������ ������� ����)
    Err := _Err;
    if not fBreak then Synchronize(Fill);
  end;
  Synchronize(Finish);
end;

procedure TTreeLoadTCP.Fill;
var
  MesText: string;
  sVersion: shortstring;
begin
  if not Assigned(Dev) then EXIT;
  if Err = NO_ERROR then
  begin
    if Dev = frmMain.GetSelectedDev(frmMain.TCPTree,0) then
    begin
      Dev.TCPFillArchTree;
    end;
    MesText := '�������� ������ ��������� ���������.';
  end
  else begin
    MesText := '�������� ������ ��������� �����������. ' + SysErrorMessage(Err) + '.';
    // ������ (��� ���-�� ������ ����) �������������, �� ������ ���� �� �����, �.�. �������� ������ ��������
    if Err <> TCP_ERROR_BAD_NET_PATH then
    begin
      sVersion := Dev.GetTCPVersion;
      if Length(sVersion) > 0 then
        MesText := MesText + ' ' + '��������! ������� ���������� ������� �������� ��� ������ � ���������, �������� ' +
                                   '������ ����������� �� �� ���� ' + sVersion + '. ������� ��������� ������ �������� ����� ' +
                                   '�� ����� www.elemer.ru'
      else
        MesText := MesText + ' ' + '��������! ������� ���������� ������� �������� ��� ������ � ���������, �������� ' +
                                   '���������� ������ ����������� ��. ������� ��������� ������ �������� ����� ' +
                                   '�� ����� www.elemer.ru';
    end;
    if Dev = frmMain.GetSelectedDev(frmMain.TCPTree,0) then
      frmArchiveTCP.memoMessage.SetTextBuf(PWideChar(MesText));
  end;
  Dev.AddToProtocol(MesText, mTCP);
  frmMain.spMessage.Caption := Dev.Name+'. '+MesText;
end;

procedure TTreeLoadTCP.Start;
var
  MesText: string;
begin
  frmMain.TCPTree.Enabled := false;
  if not Assigned(Dev) then EXIT;
  MesText := '�������� ������ ���������...';
  if Dev = frmMain.GetSelectedDev(frmMain.TCPTree,0) then
  begin
    frmArchiveTCP.DirTree.Visible := false;
    frmArchiveTCP.memoMessage.SetTextBuf(PWideChar(MesText));
    Dev.AddToProtocol(MesText, mTCP);
    frmMain.spMessage.Caption := Dev.Name+'. '+MesText;
  end;
end;

procedure TTreeLoadTCP.Finish;
begin
  frmArchiveTCP.btnRefresh.Enabled := true;
  frmMain.TCPTree.Enabled := true;
end;

procedure TFTPGetThread.Execute;
begin
  while not Terminated do if bNew then
  begin
    try
      FStream := TFileStream.Create(TargetPath, fmCreate);
      IdFTP.Get(SourcePath, FStream, bResume);
      bDone := true;
    except
      on E: EIdReadTimeout do Mess := '��������� ������������ ����� �������� ������ (' +IntToStr(Trunc(IdFTP.ReadTimeout/1000)) + ' ���)';
      on E: EIdException   do Mess := E.Message;
      on E: Exception      do Mess := E.Message;
    end;
    StreamFree;
    bNew := false;
  end;
end;

procedure TFTPGetThread.StreamFree;
begin
  FStream.Free;
end;

end.

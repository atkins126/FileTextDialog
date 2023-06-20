unit FileTextDialog;

interface

uses System.Classes, System.SysUtils, Winapi.ShlObj, Vcl.Dialogs;

type
  TFileTextDialog = class;

  TMyFileDialogEvents = class(TInterfacedObject, IFileDialogEvents,
    IFileDialogControlEvents)
  private
    FFileDialog: TFileTextDialog;
  public
    constructor Create(AFileDialog: TFileTextDialog);
    function OnFileOk(const pfd: IFileDialog): HResult; stdcall;
    function OnFolderChanging(const pfd: IFileDialog;
      const psiFolder: IShellItem): HResult; stdcall;
    function OnFolderChange(const pfd: IFileDialog): HResult; stdcall;
    function OnSelectionChange(const pfd: IFileDialog): HResult; stdcall;
    function OnShareViolation(const pfd: IFileDialog; const psi: IShellItem;
      out pResponse: Cardinal): HResult; stdcall;
    function OnTypeChange(const pfd: IFileDialog): HResult; stdcall;
    function OnOverwrite(const pfd: IFileDialog; const psi: IShellItem;
      out pResponse: Cardinal): HResult; stdcall;
    { IFileDialogControlEvents }
    function OnItemSelected(const pfdc: IFileDialogCustomize; dwIDCtl: Cardinal;
      dwIDItem: Cardinal): HResult; stdcall;
    function OnButtonClicked(const pfdc: IFileDialogCustomize;
      dwIDCtl: Cardinal): HResult; stdcall;
    function OnCheckButtonToggled(const pfdc: IFileDialogCustomize;
      dwIDCtl: Cardinal; bChecked: LONGBOOL): HResult; stdcall;
    function OnControlActivating(const pfdc: IFileDialogCustomize;
      dwIDCtl: Cardinal): HResult; stdcall;
  end;

  TFileTextDialog = class(TCustomFileDialog)
  private
    FEncodings: TStrings;
    FEncodingIndex: Integer;
    FCookie: Cardinal;
    FEvents: IFileDialogEvents;
    //FFileDialog: IFileDialog;
    function IsEncodingStored: Boolean;
    procedure SetEncodings(const Value: TStrings);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure DoTheExecute(Sender: TObject);
  published
    property Encodings: TStrings read FEncodings write SetEncodings stored IsEncodingStored;
    property EncodingIndex: Integer read FEncodingIndex write FEncodingIndex default 0;
  end;

  TFileTextOpenDialog = class(TFileTextDialog)
  strict protected
    function CreateFileDialog: IFileDialog; override;
    function GetResults: HResult; override;
  protected
    function SelectionChange: HResult; override;
  published
    property ClientGuid;
    property DefaultExtension;
    property DefaultFolder;
    property FavoriteLinks;
    property FileName;
    property FileNameLabel;
    property FileTypes;
    property FileTypeIndex;
    property OkButtonLabel;
    property Options;
    property Title;
    property OnExecute;
    property OnFileOkClick;
    property OnFolderChange;
    property OnFolderChanging;
    property OnSelectionChange;
    property OnShareViolation;
    property OnTypeChange;
  end;

  TFileTextSaveDialog = class(TFileTextDialog)
  strict protected
    function CreateFileDialog: IFileDialog; override;
  published
    property ClientGuid;
    property DefaultExtension;
    property DefaultFolder;
    property FavoriteLinks;
    property FileName;
    property FileNameLabel;
    property FileTypes;
    property FileTypeIndex;
    property OkButtonLabel;
    property Options;
    property Title;
    property OnExecute;
    property OnFileOkClick;
    property OnFolderChange;
    property OnFolderChanging;
    property OnOverwrite;
    property OnSelectionChange;
    property OnShareViolation;
    property OnTypeChange;
  end;


procedure Register;

implementation

uses Vcl.Consts, Winapi.ActiveX;

procedure Register;
begin
  RegisterComponents('GonzaloGHM', [TFileTextOpenDialog, TFileTextSaveDialog]);
end;

const
  DefaultEncodingNames: array[0..5] of string =
    (SANSIEncoding, SASCIIEncoding, SUnicodeEncoding,
     SBigEndianEncoding, SUTF8Encoding, SUTF7Encoding);

// Returns the standard encoding referenced by Name, or
// nil if Name isn't one of the encodings in DefaultEncodingNames.
function StandardEncodingFromName(const Name: string): TEncoding;
begin
  Result := nil;
  if Name = SANSIEncoding then
    Result := TEncoding.Default
  else if Name = SASCIIEncoding then
    Result := TEncoding.ASCII
  else if Name = SUnicodeEncoding then
    Result := TEncoding.Unicode
  else if Name = SBigEndianEncoding then
    Result := TEncoding.BigEndianUnicode
  else if Name = SUTF7Encoding then
    Result := TEncoding.UTF7
  else if Name = SUTF8Encoding then
    Result := TEncoding.UTF8;
end;

const
  dwGroupID: Cardinal = 1900;
  dwLabelId: Cardinal = 1901;
  dwComboId: Cardinal = 1902;
  dwFirstComboId: Cardinal = 1903;

function TMyFileDialogEvents.OnFileOk(const pfd: IFileDialog): HResult;
var
  c: IFileDialogCustomize;
  itemsel: Cardinal;
begin
  if Supports(FFileDialog.Dialog, IFileDialogCustomize, c) then
  begin
    if c.GetSelectedControlItem(dwComboID, itemsel) = S_OK then
    begin
      FFileDialog.EncodingIndex:= Integer(itemsel - dwFirstComboID);
      Result := S_OK;
      exit;
    end;
  end;
  Result:= S_FALSE;
end;

function TMyFileDialogEvents.OnFolderChange(const pfd: IFileDialog): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnFolderChanging(const pfd: IFileDialog;
  const psiFolder: IShellItem): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnOverwrite(const pfd: IFileDialog;
  const psi: IShellItem; out pResponse: Cardinal): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnSelectionChange(const pfd: IFileDialog): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnShareViolation(const pfd: IFileDialog;
  const psi: IShellItem; out pResponse: Cardinal): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnTypeChange(const pfd: IFileDialog): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnItemSelected(const pfdc: IFileDialogCustomize; dwIDCtl: Cardinal;
  dwIDItem: Cardinal): HResult;
begin
  if dwIDCtl = dwComboID then begin
    // ...
    Result := S_OK;
  end else begin
    Result := E_NOTIMPL;
  end;
end;

constructor TMyFileDialogEvents.Create(AFileDialog: TFileTextDialog);
begin
  inherited Create;
  FFileDialog := AFileDialog;
end;

function TMyFileDialogEvents.OnButtonClicked(const pfdc: IFileDialogCustomize; dwIDCtl: Cardinal): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnCheckButtonToggled(const pfdc: IFileDialogCustomize;
  dwIDCtl: Cardinal; bChecked: LONGBOOL): HResult;
begin
  Result := E_NOTIMPL;
end;

function TMyFileDialogEvents.OnControlActivating(const pfdc: IFileDialogCustomize; dwIDCtl: Cardinal): HResult;
begin
  Result := E_NOTIMPL;
end;

{ TFileTextDialog }

constructor TFileTextDialog.Create(AOwner: TComponent);
var
  I: Integer;
begin
  inherited Create(AOwner);
  FEncodings := TStringList.Create;
  for I := 0 to Length(DefaultEncodingNames) - 1 do
    FEncodings.Add(DefaultEncodingNames[I]);
  FEncodingIndex := 0;
  OnExecute:= DoTheExecute;
end;

destructor TFileTextDialog.Destroy;
begin
  if (Dialog <> nil) and (FCookie <> 0) then
    Dialog.Unadvise(FCookie);
  FEvents := nil;
  FCookie := 0;
  FEncodings.Free;
  inherited;
end;

procedure TFileTextDialog.DoTheExecute(Sender: TObject);
var
  c: IFileDialogCustomize;
  d: IFileDialogEvents;
  ck: Cardinal;
  I: Integer;
begin
  if Supports(Dialog, IFileDialogCustomize, c) then
  begin
    // Add a Advanced Button
    c.StartVisualGroup(dwGroupID, '');
    c.AddText(dwLabelID, 'Encodings');
    c.AddComboBox(dwComboID);
    for I:= 0 to FEncodings.Count -1 do
      c.AddControlItem(dwComboID, Cardinal(dwFirstComboId) + Cardinal(I),
        PWideChar(FEncodings[I]));
    c.SetSelectedControlItem(dwComboID, dwFirstComboID);
    c.EndVisualGroup;
    d := TMyFileDialogEvents.Create(Self);
    if Dialog.Advise(d, ck) = S_OK then
    begin
      FEvents := d;
      FCookie := ck;
    end;
  end;
end;

function TFileTextDialog.IsEncodingStored: Boolean;
var
  I: Integer;
begin
  Result := FEncodings.Count <> Length(DefaultEncodingNames);
  if not Result then
    for I := 0 to FEncodings.Count - 1 do
      if AnsiCompareText(FEncodings[I], DefaultEncodingNames[I]) <> 0 then
      begin
        Result := True;
        Break;
      end;
end;

procedure TFileTextDialog.SetEncodings(const Value: TStrings);
begin
  FEncodings.Assign(Value);
end;

{TFileTextOpenDialog}

function TFileTextOpenDialog.CreateFileDialog: IFileDialog;
var
  LGuid: TGUID;
begin
{$IF DEFINED(CLR)}
  LGuid := Guid.Create(CLSID_FileOpenDialog);
{$ELSE}
  LGuid := CLSID_FileOpenDialog;
{$ENDIF}
  CoCreateInstance(LGuid, nil, CLSCTX_INPROC_SERVER,
    StringToGUID(SID_IFileOpenDialog), Result);
end;

function TFileTextOpenDialog.GetResults: HResult;
var
  SItems: ^IShellItemArray;
begin
  if not (fdoAllowMultiSelect in Options) then
    Result := inherited GetResults
  else
  begin
    SItems:= @ShellItems;
    Result := (Dialog as IFileOpenDialog).GetResults(SItems^);
    if Succeeded(Result) then
      Result := GetFileNames(SItems^);
  end;
end;

function TFileTextOpenDialog.SelectionChange: HResult;
var
  SItems: ^IShellItemArray;
  SItem: ^IShellItem;
begin
  if not (fdoAllowMultiSelect in Options) then
    Result := inherited SelectionChange
  else
  begin
    SItems:= @ShellItems;
    Result := (Dialog as IFileOpenDialog).GetSelectedItems(SItems^);
    if Succeeded(Result) then
    begin
      Result := GetFileNames(SItems^);
      if Succeeded(Result) then
      begin
        SItem:= @ShellItem;
        Dialog.GetCurrentSelection(SItem^);
        DoOnSelectionChange;
      end;
      SItems^ := nil;
    end;
  end;
end;

{ TFileTextSaveDialog }

function TFileTextSaveDialog.CreateFileDialog: IFileDialog;
var
  LGuid: TGUID;
begin
{$IF DEFINED(CLR)}
  LGuid := Guid.Create(CLSID_FileSaveDialog);
{$ELSE}
  LGuid := CLSID_FileSaveDialog;
{$ENDIF}
  CoCreateInstance(LGuid, nil, CLSCTX_INPROC_SERVER,
    StringToGUID(SID_IFileSaveDialog), Result);
end;

end.

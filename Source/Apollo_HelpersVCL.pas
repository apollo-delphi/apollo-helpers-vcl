unit Apollo_HelpersVCL;

interface

uses
  System.Classes,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Graphics,
  Vcl.Grids,
  Vcl.Imaging.jpeg,
  Vcl.StdCtrls;

type
  TValidValueFunc = reference to function(const aValue: Variant; out aErrMsg: string): Boolean;
  TValidIndexFunc = reference to function(const aIndex: Integer; out aErrMsg: string): Boolean;
  TValidPassedProc = reference to procedure(const aValue: Variant);

  TWinControlHelper = class helper for TWinControl
  public
    function GetParentForm: TForm;
    procedure SetReadOnly(const aControls: TArray<TWinControl>; const aReadOnly: Boolean = True);
  end;

  TStringGridHelper = class helper for TStringGrid
  private
    procedure HideCellComboBox(Sender: TObject);
  public
    function GetColWidths: TArray<Integer>;
    function GetGridState : TGridState;
    /// recommend to call in StringGrid`s SelectCell and SetEditText event
    procedure SetCellAsComboBox(const aCol, aRow: Integer; aComboBox: TComboBox);
    procedure SetHeaders(const aColNames: TArray<string>; const aDefaultColWidths, aStoredColWidths: TArray<Integer>);
    /// recommend to call aBalloonHint.HideHint in owner form OnResize event
    procedure ValidateCell(aBalloonHint: TBalloonHint; const aCol, aRow: Integer;
      aValidFunc: TValidValueFunc; aValidPassedProc: TValidPassedProc);
  end;

  TComboBoxHelper = class helper for TComboBox
  public
    /// recommend to call aBalloonHint.HideHint in owner form OnCanResize event
    procedure Validate(aBalloonHint: TBalloonHint; aValidFunc: TValidIndexFunc;
      aValidPassedProc: TValidPassedProc);
  end;

  TCustomEditHelper = class helper for TCustomEdit
  private
    procedure FilterInputForCurrency(Sender: TObject; var Key: Char);
  public
    procedure SetFilterInputForCurrency;
    /// recommend to call aBalloonHint.HideHint in owner form OnCanResize event
    procedure Validate(aBalloonHint: TBalloonHint; aValidFunc: TValidValueFunc;
      aValidPassedProc: TValidPassedProc);
  end;

  TBalloonHintHelper = class helper for TBalloonHint
  public
    procedure ShowHintForControl(const aMsg: string; aControl: TWinControl);
  end;

  TPictureHelper = class helper for TPicture
  public
    function MakeThumbnail(const aHeight: Integer): TJpegImage;
  end;

  TVCLTools = record
    class procedure OpenURL(const aURL: string); static;
    class procedure ShowDirectory(const aDirPath: string); static;
  end;

implementation

uses
  System.SysUtils,
  System.Types,
  Vcl.ExtCtrls,
  Winapi.ActiveX,
  Winapi.GDIPOBJ,
  Winapi.ShellAPI,
  Winapi.Windows;

{ TStringGridHelper }

function TStringGridHelper.GetColWidths: TArray<Integer>;
var
  i: Integer;
begin
  Result := [];

  for i := 0 to ColCount - 1 do
    Result := Result + [ColWidths[i]];
end;

function TStringGridHelper.GetGridState: TGridState;
begin
  Result := FGridState;
end;

procedure TStringGridHelper.HideCellComboBox(Sender: TObject);
var
  ComboBox: TComboBox;
begin
  ComboBox := Sender as TComboBox;

  Cells[Col, Row] := ComboBox.Items[ComboBox.ItemIndex];
  ComboBox.Visible := False;
  SetFocus;
end;

procedure TStringGridHelper.SetCellAsComboBox(const aCol, aRow: Integer;
  aComboBox: TComboBox);
var
  CellRect: TRect;
begin
  CellRect := Self.CellRect(aCol, aRow);

  CellRect.Left := CellRect.Left + Left + 1{border};
  CellRect.Right := CellRect.Right + Left + 1{border};
  CellRect.Top := CellRect.Top + Top + 1{border};
  CellRect.Bottom := CellRect.Bottom + Top + 1{border};

  aComboBox.Left := CellRect.Left + 1;
  aComboBox.Top := CellRect.Top;
  aComboBox.Width := CellRect.Width;
  aComboBox.Height := CellRect.Height - 2;

  aComboBox.OnExit := HideCellComboBox;
  aComboBox.Style := csOwnerDrawFixed;
  aComboBox.Visible := True;
  aComboBox.SetFocus;
end;

procedure TStringGridHelper.SetHeaders(const aColNames: TArray<string>;
  const aDefaultColWidths, aStoredColWidths: TArray<Integer>);
var
  ÑhosenColWidths: TArray<Integer>;
  i: Integer;
begin
  if Length(aColNames) <> Length(aStoredColWidths) then
    ÑhosenColWidths := aDefaultColWidths
  else
  if Length(aStoredColWidths) > 0 then
    ÑhosenColWidths := aStoredColWidths
  else
    ÑhosenColWidths := aDefaultColWidths;

  RowCount := 2;
  FixedRows := 1;
  ColCount := Length(aColNames);

  for i := 0 to Length(aColNames) - 1 do
  begin
    Cells[i, 0] := aColNames[i];
    ColWidths[i] := ÑhosenColWidths[i];
  end;
end;

procedure TStringGridHelper.ValidateCell(aBalloonHint: TBalloonHint; const aCol,
  aRow: Integer; aValidFunc: TValidValueFunc; aValidPassedProc: TValidPassedProc);
var
  ErrMsg: string;
  i: Integer;
  Point: TPoint;
  Value: string;
begin
  Value := Cells[aCol, aRow];

  if aValidFunc(Value, {out}ErrMsg) then
  begin
    if Assigned(aValidPassedProc) then
      aValidPassedProc(Value)
  end
  else
  begin
    aBalloonHint.Title := 'Cell validation error';
    aBalloonHint.Description := ErrMsg;

    Point.X := 0;
    for i := 0 to aCol do
      if i = aCol then
        Point.X := Point.X + (ColWidths[i] + GridLineWidth) div 2
      else
        Point.X := Point.X + (ColWidths[i] + GridLineWidth);

    Point.Y := 0;
    for i := 0 to aRow do
      Point.Y := Point.Y + (RowHeights[i] + GridLineWidth);

    aBalloonHint.ShowHint(ClientToScreen(Point));

    Abort;
  end;
end;

{ TComboBoxHelper }

procedure TComboBoxHelper.Validate(aBalloonHint: TBalloonHint;
  aValidFunc: TValidIndexFunc; aValidPassedProc: TValidPassedProc);
var
  ErrMsg: string;
begin
  if aValidFunc(ItemIndex, {out}ErrMsg) then
  begin
    if Assigned(aValidPassedProc) then
      aValidPassedProc(ItemIndex);
  end
  else
  begin
    aBalloonHint.ShowHintForControl(ErrMsg, Self);
    Abort;
  end;
end;

{ TWinControlHelper }

function TWinControlHelper.GetParentForm: TForm;
begin
  Result := nil;

  if Assigned(Parent) then
  begin
    if Parent.InheritsFrom(TForm) then
      Exit(TForm(Parent))
    else
      Result := Parent.GetParentForm;
  end;
end;

procedure TWinControlHelper.SetReadOnly(const aControls: TArray<TWinControl>; const aReadOnly: Boolean);

  function GetControlColor: TColor;
  begin
    if aReadOnly then
      Result := cl3DLight
    else
      Result := clWindow;
  end;

var
  Control: TWinControl;
begin
  for Control in aControls do
  begin
    if Control.InheritsFrom(TLabeledEdit) then
    begin
      TLabeledEdit(Control).ReadOnly := aReadOnly;
      TLabeledEdit(Control).Color := GetControlColor;
    end
    else
    if Control.InheritsFrom(TEdit) then
    begin
      TEdit(Control).ReadOnly := aReadOnly;
      TEdit(Control).Color := GetControlColor;
    end
    else
      raise Exception.Create('TWinControlHelper.SetReadOnly: class type is not supported.');
  end;
end;

{ TCustomEditHelper }

procedure TCustomEditHelper.FilterInputForCurrency(Sender: TObject;
  var Key: Char);
begin
  if not(CharInSet(Key, ['0'..'9', #8, FormatSettings.DecimalSeparator])) then
    Key := #0;
end;

procedure TCustomEditHelper.SetFilterInputForCurrency;
begin
  Alignment := taRightJustify;
  OnKeyPress := FilterInputForCurrency;
end;

procedure TCustomEditHelper.Validate(aBalloonHint: TBalloonHint;
  aValidFunc: TValidValueFunc; aValidPassedProc: TValidPassedProc);
var
  ErrMsg: string;
begin
  if aValidFunc(Text, {out}ErrMsg) then
  begin
    if Assigned(aValidPassedProc) then
      aValidPassedProc(Text);
  end
  else
  begin
    aBalloonHint.ShowHintForControl(ErrMsg, Self);
    Abort;
  end;
end;

{ TBalloonHintHelper }

procedure TBalloonHintHelper.ShowHintForControl(const aMsg: string; aControl: TWinControl);
var
  Point: TPoint;
begin
  Title := 'Validation error';
  Description := aMsg;
  HideAfter := 3000;

  Point.X := aControl.Left + aControl.Width div 2;
  Point.Y := aControl.Top + aControl.Height;

  ShowHint(aControl.Parent.ClientToScreen(Point));
end;

{ TPictureHelper }

function TPictureHelper.MakeThumbnail(const aHeight: Integer): TJpegImage;
var
  GDIPlus: TGPGraphics;
  GPImage : TGPImage;
  MemoryStream: TMemoryStream;
  Ratio: Extended;
  Stream: IStream;
  TempBitmap: Vcl.Graphics.TBitmap;
begin
  Result := TJpegImage.Create;
  Ratio := Width / Height;

  TempBitmap := Vcl.Graphics.TBitmap.Create;
  MemoryStream := TMemoryStream.Create;
  try
    TempBitmap.PixelFormat := pf32Bit;
    TempBitmap.Height := aHeight;
    TempBitmap.Width := Trunc(aHeight * Ratio);

    SaveToStream(MemoryStream);
    MemoryStream.Position := 0;
    Stream := TStreamAdapter.Create(MemoryStream, soReference) as IStream;

    GDIPlus := TGPGraphics.Create(TempBitmap.Canvas.Handle);
    GPImage := TGPImage.Create(Stream);
    try
      GDIPlus.DrawImage(GPImage, 0, 0, TempBitmap.Width, TempBitmap.Height);
      Result.Assign(TempBitmap);
    finally
      GPImage.Free;
      GDIPlus.Free;
    end;
  finally
    TempBitmap.Free;
    MemoryStream.Free;
  end;
end;

{ TVCLTools }

class procedure TVCLTools.ShowDirectory(const aDirPath: string);
begin
  ShellExecute(Application.Handle,
    PChar('explore'),
    PChar(aDirPath),
    nil,
    nil,
    SW_SHOWNORMAL
  );
end;

class procedure TVCLTools.OpenURL(const aURL: string);
begin
  ShellExecute(Application.Handle,
    PChar('open'),
    PChar(aURL),
    nil,
    nil,
    SW_SHOWNORMAL
  );
end;

end.

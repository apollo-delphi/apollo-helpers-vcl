unit Apollo_HelpersVCL;

interface

uses
  System.Classes,
  System.Types,
  Vcl.Controls,
  Vcl.ExtCtrls,
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
    procedure FitParentHeight(const aHeight: Integer);
    procedure FitParentWidth(const aWidth: Integer);
    procedure SetReadOnly(const aControls: TArray<TWinControl>; const aReadOnly: Boolean = True);
  end;

  TFormHelper = class helper for TForm
    procedure ShowFrame(aParent: TPanel; aFrame: TFrame);
  end;

  TCellTextAlign = (ctaVertical, ctaHorizontal, ctaCenter);

  TStringGridHelper = class helper for TStringGrid
  private
    function GetTextRect(const aText: string): TRect;
    procedure DoTextOut(const aTextRect: TRect; const aText: string);
    procedure HideCellComboBox(Sender: TObject);
  public
    function GetColWidths: TArray<Integer>;
    function GetGridState : TGridState;
    function IsFixedRow(const aRow: Integer): Boolean;
    /// call in OnDrawCell event only
    function SetCellFit(const aCol, aRow: Integer; const aTextAlign: TCellTextAlign;
      aBackgoundColor: TColor = clNone): TRect;
    procedure DeleteRow(const aRow: Integer);
    procedure InsertFirstRow;
    procedure InsertRow(const aBeforeRow: Integer);
    /// recommend to call in StringGrid`s SelectCell and SetEditText event
    procedure SetCellAsComboBox(const aCol, aRow: Integer; aComboBox: TComboBox);
    procedure SetHeaders(const aColNames: TArray<string>; const aDefaultColWidths, aStoredColWidths: TArray<Integer>);
    /// recommend to call aBalloonHint.HideHint in owner form OnCanResize event
    procedure ValidateCell(aBalloonHint: TBalloonHint; const aCol, aRow: Integer;
      aValidFunc: TValidValueFunc; aValidPassedProc: TValidPassedProc);
  end;

  TComboBoxHelper = class helper for TComboBox
  type
    TForEachObjectProc<T: class> = reference to procedure(aObject: T; const aIndex: Integer);
  public
    function GetSelectedObject<T: class>: T;
    procedure ForEachObject<T: class>(aForEachObjectProc: TForEachObjectProc<T>);
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
    function ResizeByHeight(const aHeight: Integer): TRect;
    procedure LoadFromResource(aInstance: HInst; const aResourceName: string);
  end;

  TControlsArrayHelper = record helper for TArray<TControl>
    function Contains(aControl: TControl): Boolean;
  end;

  TVCLTools = record
    class procedure OpenURL(const aURL: string); static;
    class procedure ShowDirectory(const aDirPath: string); static;
    class procedure ShowFrame(aParent: TPanel; aFrame: TFrame); static;
  end;

implementation

uses
  System.Math,
  System.SysUtils,
  Winapi.ActiveX,
  Winapi.GDIPOBJ,
  Winapi.ShellAPI,
  Winapi.Windows;

{ TStringGridHelper }

function TStringGridHelper.SetCellFit(const aCol, aRow: Integer; const aTextAlign: TCellTextAlign;
  aBackgoundColor: TColor): TRect;
var
  CellRec: TRect;
  Text: string;
begin
  CellRec := CellRect(aCol, aRow);
  Text := Cells[aCol, aRow];

  Result := GetTextRect(Text);
  Result.SetLocation(CellRec.Left + 2, CellRec.Top + 2);

  case aTextAlign of
    ctaVertical: Result.AlignVertical(CellRec);
    ctaHorizontal: Result.AlignHorizontal(CellRec);
    ctaCenter:
      begin
        Result.AlignHorizontal(CellRec);
        Result.AlignVertical(CellRec);
      end;
  end;

  if aBackgoundColor = clNone then
  begin
    if aRow < FixedRows then
      Canvas.Brush.Color := FixedColor
    else
      Canvas.Brush.Color := Color;
  end
  else
    Canvas.Brush.Color := aBackgoundColor;

  Canvas.FillRect(CellRec);
  DoTextOut(Result, Text);
end;

procedure TStringGridHelper.DeleteRow(const aRow: Integer);
var
  Col: Integer;
  Row: Integer;
begin
  if (aRow = FixedRows) and (RowCount = FixedRows + 1) then
    for Col := 0 to ColCount - 1 do
    begin
      Cells[Col, aRow] := '';
      Objects[Col, aRow] := nil;
    end
  else
  begin
    for Row := aRow + 1 to RowCount - 1 do
      for Col := 0 to ColCount - 1 do
      begin
        Cells[Col, Row - 1] := Cells[Col, Row];
        Objects[Col, Row - 1] := Objects[Col, Row];
      end;

    RowCount := RowCount - 1;
  end;
end;

procedure TStringGridHelper.DoTextOut(const aTextRect: TRect; const aText: string);
begin
  Canvas.TextOut(aTextRect.Left, aTextRect.Top, aText);
end;

function TStringGridHelper.GetColWidths: TArray<Integer>;
var
  i: Integer;
begin
  Result := [];

  for i := 0 to ColCount - 1 do
    Result := Result + [ColWidths[i]];
end;

function TStringGridHelper.IsFixedRow(const aRow: Integer): Boolean;
begin
  Result := aRow < FixedRows;
end;

function TStringGridHelper.GetGridState: TGridState;
begin
  Result := FGridState;
end;

function TStringGridHelper.GetTextRect(const aText: string): TRect;
begin
  Result.Top := 0;
  Result.Left := 0;
  Result.Height := Canvas.TextHeight(aText);
  Result.Width := Canvas.TextWidth(aText);
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

procedure TStringGridHelper.InsertFirstRow;
begin
  if RowCount > FixedRows + 1 then
    InsertRow(0);
end;

procedure TStringGridHelper.InsertRow(const aBeforeRow: Integer);
var
  Col: Integer;
  FirstRow: Integer;
  Row: Integer;
begin
  RowCount := RowCount + 1;

  FirstRow := Max(FixedRows, aBeforeRow);

  for Row := RowCount - 2 downto FirstRow do
    for Col := 0 to ColCount - 1 do
    begin
      Cells[Col, Row + 1] := Cells[Col, Row];
      Objects[Col, Row + 1] := Objects[Col, Row];
    end;

  for Col := 0 to ColCount - 1 do
  begin
    Cells[Col, FirstRow] := '';
    Objects[Col, FirstRow] := nil;
  end;
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
    aBalloonHint.HideAfter := 3000;

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

function TComboBoxHelper.GetSelectedObject<T>: T;
begin
  Result := Items.Objects[ItemIndex] as T;
end;

procedure TComboBoxHelper.ForEachObject<T>(aForEachObjectProc: TForEachObjectProc<T>);
var
  i: Integer;
begin
  for i := 0 to Items.Count - 1 do
    aForEachObjectProc(T(Items.Objects[i]), i);
end;

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

procedure TWinControlHelper.FitParentHeight(const aHeight: Integer);
begin
  if not Assigned(Parent) then
    Exit;

  if Parent.ClientHeight < aHeight then
    Parent.ClientHeight := aHeight;

  Parent.FitParentHeight(aHeight);

end;

procedure TWinControlHelper.FitParentWidth(const aWidth: Integer);
begin
  if not Assigned(Parent) then
    Exit;

  if Parent.ClientWidth < aWidth then
    Parent.ClientWidth := aWidth;

  Parent.FitParentWidth(aWidth);
end;

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
    if Control.InheritsFrom(TComboBox) then
    begin
      TComboBox(Control).Enabled := not aReadOnly;
      TComboBox(Control).Color := GetControlColor;
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

procedure TPictureHelper.LoadFromResource(aInstance: HInst; const aResourceName: string);
var
  ResStream: TResourceStream;
begin
  ResStream := TResourceStream.Create(aInstance, aResourceName, RT_RCDATA);
  try
    LoadFromStream(ResStream);
  finally
    ResStream.Free;
  end;
end;

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

function TPictureHelper.ResizeByHeight(const aHeight: Integer): TRect;
var
  Ratio: Extended;
begin
  Result.Top := 0;
  Result.Left := 0;
  Result.Height := aHeight;

  Ratio := aHeight / Height;
  Result.Width := Trunc(Width * Ratio);
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

class procedure TVCLTools.ShowFrame(aParent: TPanel; aFrame: TFrame);
var
  Frame: TFrame;
  i: Integer;
begin
  for i := 0 to aParent.ControlCount - 1 do
    if aParent.Controls[i].InheritsFrom(TFrame) then
    begin
      Frame := TFrame(aParent.Controls[i]);

      if Frame <> aFrame then
        Frame.Visible := False;
    end;

  aFrame.Parent := aParent;
  aFrame.Align := alClient;
  aFrame.FitParentHeight(aFrame.Constraints.MinHeight);
  aFrame.FitParentWidth(aFrame.Constraints.MinWidth);

  aFrame.Visible := True;
end;

{ TFormHelper }

procedure TFormHelper.ShowFrame(aParent: TPanel; aFrame: TFrame);
begin
  TVCLTools.ShowFrame(aParent, aFrame);
end;

{TControlsArrayHelper}

function TControlsArrayHelper.Contains(aControl: TControl): Boolean;
var
  Control: TControl;
begin
  Result := False;

  for Control in Self do
    if Control = aControl then
      Exit(True);
end;

end.

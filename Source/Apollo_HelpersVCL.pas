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
  TVCLTools = record
    class function DlgConfirm(const aMsg: string): Boolean; static;
    class function MakeThumbnailAsStream(aPictureStream: TStream; const aHeight: Integer): TMemoryStream; static;
    class procedure ShowFrame(aParent: TPanel; aFrame: TFrame); static;
  end;

  TValidValueFunc = reference to function(const aValue: Variant; out aErrMsg: string): Boolean;
  TValidPassedProc = reference to procedure(const aValue: Variant);

  TWinControlHelper = class helper for TWinControl
  public
    procedure SetReadOnly(const aControls: TArray<TWinControl>; const aReadOnly: Boolean = True);
  end;

  TStringGridHelper = class helper for TStringGrid
  private
    function GridIsEmpty: Boolean;
  public
    procedure DeleteRow(const aRow: Integer);
    procedure InsertFirstRow;
    procedure InsertRow(const aBeforeRow: Integer);
  end;

  TCustomEditHelper = class helper for TCustomEdit
  public
    /// recommend to call aBalloonHint.HideHint in owner form OnCanResize event
    procedure Validate(aBalloonHint: TBalloonHint; aValidFunc: TValidValueFunc;
      aValidPassedProc: TValidPassedProc);
  end;

  TBalloonHintHelper = class helper for TBalloonHint
  public
    procedure ShowHintForControl(const aMsg: string; aControl: TWinControl);
  end;

  TFormHelper = class helper for TForm
    procedure ShowFrame(aParent: TPanel; aFrame: TFrame);
  end;

  TPictureHelper = class helper for TPicture
  public
    function MakeThumbnail(const aHeight: Integer): TJpegImage;
  end;

  TRectHelper = record helper for TRect
  public
    procedure AlignHorizontal(const aMasterRect: TRect);
    procedure AlignVertical(const aMasterRect: TRect);
  end;

  TCanvasHelper = class helper for TCanvas
  public
    function GetTextRect(const aText: string): TRect;
  end;

implementation

uses
  System.Math,
  System.SysUtils,
  System.UITypes,
  Vcl.Dialogs,
  Winapi.ActiveX,
  Winapi.GDIPOBJ;

{ TStringGridHelper }

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

function TStringGridHelper.GridIsEmpty: Boolean;

  function LastRowIsEmpty: Boolean;
  var
    i: Integer;
  begin
    Result := True;

    for i := 0 to ColCount - 1 do
      if Cells[i, RowCount - 1] <> '' then
        Exit(False);
  end;

begin
  if (FixedRows = RowCount - 1) and LastRowIsEmpty then
    Result := True
  else
    Result := False;
end;

procedure TStringGridHelper.InsertFirstRow;
begin
  if not GridIsEmpty then
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

{ TCustomEditHelper }

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

procedure TBalloonHintHelper.ShowHintForControl(const aMsg: string;
  aControl: TWinControl);
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

{ TWinControlHelper }

procedure TWinControlHelper.SetReadOnly(const aControls: TArray<TWinControl>;
  const aReadOnly: Boolean);

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

{ TVCLTools }

class function TVCLTools.DlgConfirm(const aMsg: string): Boolean;
begin
  if MessageDlg(aMsg, mtConfirmation,
    [mbYes, mbCancel], 0) <> mrYes
  then
    Result := False
  else
    Result := True;
end;

class function TVCLTools.MakeThumbnailAsStream(
  aPictureStream: TStream; const aHeight: Integer): TMemoryStream;
var
  Pic: TPicture;
  Jpeg: TJpegImage;
begin
  Pic := TPicture.Create;
  Result := TMemoryStream.Create;
  try
    Pic.LoadFromStream(aPictureStream);
    Jpeg := Pic.MakeThumbnail(aHeight);
    try
      Jpeg.SaveToStream(Result);
    finally
      Jpeg.Free;
    end;
  finally
    Pic.Free;
  end;
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

  aFrame.Visible := True;
end;

{ TFormHelper }

procedure TFormHelper.ShowFrame(aParent: TPanel; aFrame: TFrame);
begin
  TVCLTools.ShowFrame(aParent, aFrame);
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

{ TRectHelper }

procedure TRectHelper.AlignHorizontal(const aMasterRect: TRect);
var
  AlignedLeft: Integer;
  NewTop: Integer;
begin
  AlignedLeft := aMasterRect.Left + ((aMasterRect.Width - Width) div 2);

  if aMasterRect.Contains(TPoint.Create(AlignedLeft, Top)) then
    NewTop := Top
  else
    NewTop := aMasterRect.Top;

  SetLocation(AlignedLeft, NewTop);
end;

procedure TRectHelper.AlignVertical(const aMasterRect: TRect);
var
  AlignedTop: Integer;
  NewLeft: Integer;
begin
  AlignedTop := aMasterRect.Top + ((aMasterRect.Height - Height) div 2);

  if aMasterRect.Contains(TPoint.Create(Left, AlignedTop)) then
    NewLeft := Left
  else
    NewLeft := aMasterRect.Left;

  SetLocation(NewLeft, AlignedTop);
end;

{ TCanvasHelper }

function TCanvasHelper.GetTextRect(const aText: string): TRect;
begin
  Result.Top := 0;
  Result.Left := 0;
  Result.Height := TextHeight(aText);
  Result.Width := TextWidth(aText);
end;

end.

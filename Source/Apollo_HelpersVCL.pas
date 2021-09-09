unit Apollo_HelpersVCL;

interface

uses
  System.Classes,
  Vcl.Controls,
  Vcl.Grids,
  Vcl.StdCtrls;

type
  TValidFunc = reference to function(const aValue: Variant; out aErrMsg: string): Boolean;
  TValidPassedProc = reference to procedure(const aValue: Variant);

  TStringGridHelper = class helper for TStringGrid
  private
    procedure HideCellComboBox(Sender: TObject);
  public
    function GetGridState : TGridState;
    /// recommend to call in StringGrid`s SelectCell event
    procedure SetCellAsComboBox(const aCol, aRow: Integer; aComboBox: TComboBox);
    /// recommend to call aBalloonHint.HideHint in owner form OnResize event
    procedure ValidateCell(aBalloonHint: TBalloonHint; const aCol, aRow: Integer;
      aValidFunc: TValidFunc; aValidPassedProc: TValidPassedProc);
  end;

implementation

uses
  System.SysUtils,
  System.Types;

{ TStringGridHelper }

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

procedure TStringGridHelper.ValidateCell(aBalloonHint: TBalloonHint; const aCol,
  aRow: Integer; aValidFunc: TValidFunc; aValidPassedProc: TValidPassedProc);
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

end.

// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_returning_record
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TPoint = record X, Y: Integer; end;
type TPointArray = class
  private FPoints: array[0..1] of TPoint;
  private function GetPoint(i: Integer): TPoint;
  private procedure SetPoint(i: Integer; p: TPoint);
  public property Points[i: Integer]: TPoint read GetPoint write SetPoint; default;
end;
function TPointArray.GetPoint(i: Integer): TPoint; begin Result := FPoints[i]; end;
procedure TPointArray.SetPoint(i: Integer; p: TPoint); begin FPoints[i] := p; end;
var pa: TPointArray; pt: TPoint;
begin
  pa := TPointArray.Create;
  pt.X := 10; pt.Y := 20;
  pa[0] := pt;
  WriteLn(pa[0].X);
  WriteLn(pa[0].Y);
  pa.Free;
end.

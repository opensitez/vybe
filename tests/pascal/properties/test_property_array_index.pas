// vybe-test: pascal/properties/test_property_array_index
// origin: languages/pascal/tests/pascal/test_properties.rs
program T;
{$mode delphi}
uses SysUtils;
type
  TVector = class
  private
    FData: array[0..4] of Integer;
    function GetItem(idx: Integer): Integer;
    procedure SetItem(idx, val: Integer);
  public
    property Items[idx: Integer]: Integer read GetItem write SetItem;
  end;

function TVector.GetItem(idx: Integer): Integer;
begin
  Result := FData[idx];
end;

procedure TVector.SetItem(idx, val: Integer);
begin
  FData[idx] := val;
end;

var
  v: TVector;
  i: Integer;
begin
  v := TVector.Create;
  for i := 0 to 4 do
    v.Items[i] := i * 10;
  for i := 0 to 4 do
    WriteLn(v.Items[i]);
end.

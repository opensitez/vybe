// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_modifying_internal_counter
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TTrackedArray = class
  private FData: array[0..2] of Integer; FMutations: Integer;
  private procedure SetVal(i, val: Integer);
  private function GetVal(i: Integer): Integer;
  public constructor Create;
  public property Items[i: Integer]: Integer read GetVal write SetVal; default;
  public property Mutations: Integer read FMutations;
end;
constructor TTrackedArray.Create; begin FMutations := 0; end;
procedure TTrackedArray.SetVal(i, val: Integer); begin FData[i] := val; Inc(FMutations); end;
function TTrackedArray.GetVal(i: Integer): Integer; begin Result := FData[i]; end;
var ta: TTrackedArray;
begin
  ta := TTrackedArray.Create;
  ta[0] := 5; ta[1] := 10;
  WriteLn(ta.Mutations);
  ta.Free;
end.

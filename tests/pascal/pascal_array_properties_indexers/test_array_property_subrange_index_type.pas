// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_subrange_index_type
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TIndex = 1..3;
type TSubrangeIndexed = class
  private FData: array[TIndex] of String;
  private function GetVal(idx: TIndex): String;
  private procedure SetVal(idx: TIndex; val: String);
  public property Values[idx: TIndex]: String read GetVal write SetVal; default;
end;
function TSubrangeIndexed.GetVal(idx: TIndex): String; begin Result := FData[idx]; end;
procedure TSubrangeIndexed.SetVal(idx: TIndex; val: String); begin FData[idx] := val; end;
var si: TSubrangeIndexed;
begin
  si := TSubrangeIndexed.Create;
  si[1] := 'One';
  si[3] := 'Three';
  WriteLn(si[1]);
  WriteLn(si[3]);
  si.Free;
end.

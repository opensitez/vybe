// vybe-test: pascal/pascal_array_properties_indexers/test_array_property_explicit_name_access
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TStrList = class
  private FItems: array[0..2] of String;
  private function GetItem(i: Integer): String;
  private procedure SetItem(i: Integer; s: String);
  public property Strings[i: Integer]: String read GetItem write SetItem;
end;
function TStrList.GetItem(i: Integer): String; begin Result := FItems[i]; end;
procedure TStrList.SetItem(i: Integer; s: String); begin FItems[i] := s; end;
var list: TStrList;
begin
  list := TStrList.Create;
  list.Strings[0] := 'First';
  WriteLn(list.Strings[0]);
  list.Free;
end.

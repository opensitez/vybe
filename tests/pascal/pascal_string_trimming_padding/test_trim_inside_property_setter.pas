// vybe-test: pascal/pascal_string_trimming_padding/test_trim_inside_property_setter
// origin: languages/pascal/tests/pascal/test_pascal_string_trimming_padding.rs
program Test;
{$mode delphi}
uses SysUtils;
type TCleanItem = class
  private FName: String;
  private procedure SetName(v: String);
  public property Name: String read FName write SetName;
end;
procedure TCleanItem.SetName(v: String); begin FName := Trim(v); end;
var item: TCleanItem;
begin
  item := TCleanItem.Create;
  item.Name := '   Padded   ';
  WriteLn(item.Name);
  item.Free;
end.

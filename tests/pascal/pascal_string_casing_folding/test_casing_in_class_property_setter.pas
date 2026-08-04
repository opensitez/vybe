// vybe-test: pascal/pascal_string_casing_folding/test_casing_in_class_property_setter
// origin: languages/pascal/tests/pascal/test_pascal_string_casing_folding.rs
program Test;
{$mode delphi}
uses SysUtils;
type TUpperHolder = class
  private FTitle: String;
  private procedure SetTitle(t: String);
  public property Title: String read FTitle write SetTitle;
end;
procedure TUpperHolder.SetTitle(t: String); begin FTitle := UpperCase(t); end;
var h: TUpperHolder;
begin
  h := TUpperHolder.Create;
  h.Title := 'lowercase title';
  WriteLn(h.Title);
  h.Free;
end.

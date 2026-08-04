// vybe-test: pascal/pascal_class_properties_accessors/test_property_string_trim_on_set
// origin: languages/pascal/tests/pascal/test_pascal_class_properties_accessors.rs
program Test;
{$mode delphi}
uses SysUtils;
type TCleanText = class
  private FText: String;
  private procedure SetText(v: String);
  public property Text: String read FText write SetText;
end;
procedure TCleanText.SetText(v: String);
begin
  FText := Trim(v);
end;
var ct: TCleanText;
begin
  ct := TCleanText.Create;
  ct.Text := '   Cleaned   ';
  WriteLn('[' + ct.Text + ']');
  ct.Free;
end.

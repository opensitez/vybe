// vybe-test: pascal/pascal_string_builder_utility/test_stringbuilder_clear_reset
// origin: languages/pascal/tests/pascal/test_pascal_string_builder_utility.rs
program Test;
{$mode delphi}
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('Populated');
  sb.Clear;
  WriteLn(sb.Length);
  sb.Append('FreshData');
  WriteLn(sb.ToString);
  sb.Free;
end.

// vybe-test: pascal/pascal_string_builder_utility/test_stringbuilder_parameter_passing
// origin: languages/pascal/tests/pascal/test_pascal_string_builder_utility.rs
program Test;
{$mode delphi}
uses SysUtils;
procedure BuildHeader(sb: TStringBuilder);
begin
  sb.Append('=== HEADER ===');
end;
var b: TStringBuilder;
begin
  b := TStringBuilder.Create;
  BuildHeader(b);
  WriteLn(b.ToString);
  b.Free;
end.

// vybe-test: pascal/pascal_string_builder_utility/test_stringbuilder_basic_append
// origin: languages/pascal/tests/pascal/test_pascal_string_builder_utility.rs
program Test;
{$mode delphi}
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('Hello ');
  sb.Append('World');
  WriteLn(sb.ToString);
  sb.Free;
end.

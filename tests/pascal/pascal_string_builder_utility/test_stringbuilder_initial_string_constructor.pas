// vybe-test: pascal/pascal_string_builder_utility/test_stringbuilder_initial_string_constructor
// origin: languages/pascal/tests/pascal/test_pascal_string_builder_utility.rs
program Test;
{$mode delphi}
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create('InitialContent');
  sb.Append('_Appended');
  WriteLn(sb.ToString);
  sb.Free;
end.

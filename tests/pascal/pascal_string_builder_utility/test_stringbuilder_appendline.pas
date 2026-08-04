// vybe-test: pascal/pascal_string_builder_utility/test_stringbuilder_appendline
// origin: languages/pascal/tests/pascal/test_pascal_string_builder_utility.rs
program Test;
{$mode delphi}
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.AppendLine('Line1');
  sb.Append('Line2');
  WriteLn(sb.ToString);
  sb.Free;
end.

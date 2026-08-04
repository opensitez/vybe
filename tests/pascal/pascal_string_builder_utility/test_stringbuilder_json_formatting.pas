// vybe-test: pascal/pascal_string_builder_utility/test_stringbuilder_json_formatting
// origin: languages/pascal/tests/pascal/test_pascal_string_builder_utility.rs
program Test;
{$mode delphi}
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('{');
  sb.Append('"id":').Append(10).Append(',');
  sb.Append('"active":').Append('true');
  sb.Append('}');
  WriteLn(sb.ToString);
  sb.Free;
end.

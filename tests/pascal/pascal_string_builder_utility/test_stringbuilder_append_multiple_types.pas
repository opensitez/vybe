// vybe-test: pascal/pascal_string_builder_utility/test_stringbuilder_append_multiple_types
// origin: languages/pascal/tests/pascal/test_pascal_string_builder_utility.rs
program Test;
{$mode delphi}
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('Int: ');
  sb.Append(42);
  sb.Append(' Bool: ');
  sb.Append(True);
  WriteLn(sb.ToString);
  sb.Free;
end.

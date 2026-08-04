// vybe-test: pascal/pascal_string_builder_utility/test_stringbuilder_loop_aggregation
// origin: languages/pascal/tests/pascal/test_pascal_string_builder_utility.rs
program Test;
{$mode delphi}
uses SysUtils;
var sb: TStringBuilder; i: Integer;
begin
  sb := TStringBuilder.Create;
  for i := 1 to 5 do
  begin
    if i > 1 then sb.Append(',');
    sb.Append(i);
  end;
  WriteLn(sb.ToString);
  sb.Free;
end.

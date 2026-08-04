// vybe-test: pascal/pascal_string_builder_utility/test_stringbuilder_xml_tag_building
// origin: languages/pascal/tests/pascal/test_pascal_string_builder_utility.rs
program Test;
{$mode delphi}
uses SysUtils;
var sb: TStringBuilder;
begin
  sb := TStringBuilder.Create;
  sb.Append('<title>').Append('Pascal Documentation').Append('</title>');
  WriteLn(sb.ToString);
  sb.Free;
end.

// vybe-test: pascal/closures_extended/nested_capture_string_concat_loop
// origin: languages/pascal/tests/pascal/test_closures_extended.rs
program T;
{$mode delphi}
uses SysUtils;
procedure Outer;
var acc: String;
  procedure Append(s: String);
  begin acc := acc + s; end;
begin
  acc := '';
  Append('a'); Append('b'); Append('c');
  WriteLn(acc);
end;
begin Outer; end.

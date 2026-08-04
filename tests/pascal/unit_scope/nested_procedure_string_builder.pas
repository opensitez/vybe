// vybe-test: pascal/unit_scope/nested_procedure_string_builder
// origin: languages/pascal/tests/pascal/test_unit_scope.rs
program T;
{$mode delphi}
uses SysUtils; procedure P; var s: string; procedure Append(c: Char); begin s:=s+c; end; begin s:=''; Append('a'); Append('b'); WriteLn(s); end; begin P; end.

// vybe-test: pascal/print_extended/write_array_element_formatted
// origin: languages/pascal/tests/pascal/test_print_extended.rs
program T;
{$mode delphi}
uses SysUtils; var a: array[0..2] of Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; WriteLn(a[1]:3); end.

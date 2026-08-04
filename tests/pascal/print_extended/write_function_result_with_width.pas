// vybe-test: pascal/print_extended/write_function_result_with_width
// origin: languages/pascal/tests/pascal/test_print_extended.rs
program T;
{$mode delphi}
uses SysUtils; function N: Integer; begin Result := 88; end; begin WriteLn(N:4); end.

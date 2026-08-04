// vybe-test: pascal/print_extended/write_columns_in_procedure
// origin: languages/pascal/tests/pascal/test_print_extended.rs
program T;
{$mode delphi}
uses SysUtils; procedure Col(a, b: Integer); begin Write(a:4); WriteLn(b:4); end; begin Col(1, 20); Col(300, 4); end.

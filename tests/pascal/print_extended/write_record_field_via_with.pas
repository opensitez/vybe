// vybe-test: pascal/print_extended/write_record_field_via_with
// origin: languages/pascal/tests/pascal/test_print_extended.rs
program T;
{$mode delphi}
uses SysUtils; type TR = record V: Integer; end; var r: TR; begin r.V := 15; with r do WriteLn(V:4); end.

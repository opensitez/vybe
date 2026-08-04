// vybe-test: pascal/pascal_forward_external_declarations/test_forward_const_parameter
// origin: languages/pascal/tests/pascal/test_pascal_forward_external_declarations.rs
program Test;
{$mode delphi}
uses SysUtils;
function FormatMsg(const msg: String): String; forward;
begin
  WriteLn(FormatMsg('Alert'));
end;
function FormatMsg(const msg: String): String;
begin
  Result := '[MSG]: ' + msg;
end;

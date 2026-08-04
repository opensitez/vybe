// vybe-test: pascal/scope_advanced/test_scope_multiple_procs_share_global
// origin: languages/pascal/tests/pascal/test_scope_advanced.rs
program T;
{$mode delphi}
uses SysUtils;
var
  log: string;

procedure Append(s: string);
begin
  if log = '' then log := s
  else log := log + ',' + s;
end;

procedure Step1;
begin
  Append('step1');
end;

procedure Step2;
begin
  Append('step2');
end;

procedure Step3;
begin
  Append('step3');
end;

begin
  log := '';
  Step1;
  Step2;
  Step3;
  WriteLn(log);
end.

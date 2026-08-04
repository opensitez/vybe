// vybe-test: pascal/lambdas_anonymous/anonymous_nested_call
// origin: languages/pascal/tests/pascal/test_lambdas_anonymous.rs
program T;
{$mode delphi}
uses SysUtils; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*2; end, 4)); end.

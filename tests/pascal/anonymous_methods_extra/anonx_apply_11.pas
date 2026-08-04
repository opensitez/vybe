// vybe-test: pascal/anonymous_methods_extra/anonx_apply_11
// origin: languages/pascal/tests/pascal/test_anonymous_methods_extra.rs
program T;
{$mode delphi}
uses SysUtils; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*11; end,11)); end.

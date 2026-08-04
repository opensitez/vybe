// vybe-test: pascal/lambdas_anonymous/anonymous_method_as_callback_filter
// origin: languages/pascal/tests/pascal/test_lambdas_anonymous.rs
program T;
{$mode delphi}
uses SysUtils; function CountIf(const a:array of Integer; pred:function(n:Integer):Boolean):Integer; var i:Integer; begin Result:=0; for i:=Low(a) to High(a) do if pred(a[i]) then Inc(Result); end; begin WriteLn(CountIf([1,2,3,4], function(n:Integer):Boolean begin Result:=n mod 2=0; end)); end.

// vybe-test: pascal/delegates_events/event_compare_delegate_as_parameter
// origin: languages/pascal/tests/pascal/test_delegates_events.rs
program T;
{$mode delphi}
uses SysUtils; function Pick(a,b:Integer; cmp:function(x,y:Integer):Boolean):Integer; begin if cmp(a,b) then Result:=a else Result:=b; end; begin WriteLn(Pick(3,7, function(x,y:Integer):Boolean begin Result:=x>y; end)); end.

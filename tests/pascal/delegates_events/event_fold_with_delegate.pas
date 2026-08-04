// vybe-test: pascal/delegates_events/event_fold_with_delegate
// origin: languages/pascal/tests/pascal/test_delegates_events.rs
program T;
{$mode delphi}
uses SysUtils; function Fold(start:Integer; fn:function(acc,x:Integer):Integer):Integer; begin Result:=fn(start,0); end; begin WriteLn(Fold(10, function(acc,x:Integer):Integer begin Result:=acc+5; end)); end.

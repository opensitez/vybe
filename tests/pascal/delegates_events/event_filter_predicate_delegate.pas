// vybe-test: pascal/delegates_events/event_filter_predicate_delegate
// origin: languages/pascal/tests/pascal/test_delegates_events.rs
program T;
{$mode delphi}
uses SysUtils; function Keep(n:Integer; pred:function(x:Integer):Boolean):Boolean; begin Result:=pred(n); end; begin WriteLn(Keep(4, function(x:Integer):Boolean begin Result:=x mod 2=0; end)); end.

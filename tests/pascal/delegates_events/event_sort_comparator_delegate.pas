// vybe-test: pascal/delegates_events/event_sort_comparator_delegate
// origin: languages/pascal/tests/pascal/test_delegates_events.rs
program T;
{$mode delphi}
uses SysUtils; function Earlier(a,b:Integer; less:function(x,y:Integer):Boolean):Boolean; begin Result:=less(a,b); end; begin WriteLn(Earlier(2,5, function(x,y:Integer):Boolean begin Result:=x<y; end)); end.

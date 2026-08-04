// vybe-test: pascal/delegates_events/event_loop_break_on_delegate_result
// origin: languages/pascal/tests/pascal/test_delegates_events.rs
program T;
{$mode delphi}
uses SysUtils; function FindFirst(pred:function(i:Integer):Boolean):Integer; var i:Integer; begin Result:=-1; for i:=1 to 5 do if pred(i) then begin Result:=i; Exit; end; end; begin WriteLn(FindFirst(function(i:Integer):Boolean begin Result:=i=3; end)); end.

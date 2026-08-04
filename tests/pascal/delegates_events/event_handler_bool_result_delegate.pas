// vybe-test: pascal/delegates_events/event_handler_bool_result_delegate
// origin: languages/pascal/tests/pascal/test_delegates_events.rs
program T;
{$mode delphi}
uses SysUtils; function Test(n:Integer; ok:function(x:Integer):Boolean):Boolean; begin Result:=ok(n); end; begin WriteLn(Test(10, function(x:Integer):Boolean begin Result:=x>5; end)); end.

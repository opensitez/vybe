// vybe-test: pascal/time_calendar/encode_time_basic
// origin: languages/pascal/tests/pascal/test_time_calendar.rs
program T;
{$mode delphi}
uses SysUtils; var t:TDateTime; begin t:=EncodeTime(1,2,3,0); WriteLn(FormatDateTime("hh:nn:ss",t)); end.

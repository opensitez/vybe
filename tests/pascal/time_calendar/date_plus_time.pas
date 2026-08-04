// vybe-test: pascal/time_calendar/date_plus_time
// origin: languages/pascal/tests/pascal/test_time_calendar.rs
program T;
{$mode delphi}
uses SysUtils; var d,t,dt:TDateTime; begin d:=EncodeDate(2020,1,1); t:=EncodeTime(12,0,0,0); dt:=d+t; WriteLn(FormatDateTime("yyyy-mm-dd",dt)); WriteLn(FormatDateTime("hh:nn",dt)); end.

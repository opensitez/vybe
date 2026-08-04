// vybe-test: pascal/pascal_untyped_parameters/test_untyped_parameter_in_class_method
// origin: languages/pascal/tests/pascal/test_pascal_untyped_parameters.rs
program Test;
{$mode delphi}
uses SysUtils;
type TStreamer = class
  public procedure WriteRaw(const data; size: Integer);
end;
procedure TStreamer.WriteRaw(const data; size: Integer);
var pb: PByte;
begin
  pb := @data;
  WriteLn(pb^);
end;
var s: TStreamer; num: Integer;
begin
  num := 42;
  s := TStreamer.Create;
  s.WriteRaw(num, SizeOf(Integer));
  s.Free;
end.

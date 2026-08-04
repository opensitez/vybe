// vybe-test: pascal/pascal_interface_reference_counting/test_multiple_interfaces_on_single_class
// origin: languages/pascal/tests/pascal/test_pascal_interface_reference_counting.rs
program Test;
{$mode delphi}
uses SysUtils;
type IReader = interface
  ['{33333333-3333-3333-3333-333333333333}']
  function ReadData: String;
end;
type IWriter = interface
  ['{44444444-4444-4444-4444-444444444444}']
  procedure WriteData(s: String);
end;
type TFileHandler = class(TInterfacedObject, IReader, IWriter)
  private FBuffer: String;
  public function ReadData: String;
  public procedure WriteData(s: String);
end;
function TFileHandler.ReadData: String; begin Result := FBuffer; end;
procedure TFileHandler.WriteData(s: String); begin FBuffer := s; end;
var r: IReader; w: IWriter; h: TFileHandler;
begin
  h := TFileHandler.Create;
  w := h; r := h;
  w.WriteData('StreamContent');
  WriteLn(r.ReadData);
end.

// vybe-test: pascal/pascal_array_properties_indexers/test_write_only_array_property
// origin: languages/pascal/tests/pascal/test_pascal_array_properties_indexers.rs
program Test;
{$mode delphi}
uses SysUtils;
type TBufferWriter = class
  private FBuf: array[0..2] of Integer;
  private procedure SetBuf(i: Integer; val: Integer);
  public property Buffer[i: Integer]: Integer write SetBuf;
  public function GetSum: Integer;
end;
procedure TBufferWriter.SetBuf(i: Integer; val: Integer); begin FBuf[i] := val; end;
function TBufferWriter.GetSum: Integer; begin Result := FBuf[0] + FBuf[1] + FBuf[2]; end;
var bw: TBufferWriter;
begin
  bw := TBufferWriter.Create;
  bw.Buffer[0] := 10;
  bw.Buffer[1] := 20;
  bw.Buffer[2] := 30;
  WriteLn(bw.GetSum);
  bw.Free;
end.

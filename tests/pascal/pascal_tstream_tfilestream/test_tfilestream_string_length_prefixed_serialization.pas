// vybe-test: pascal/pascal_tstream_tfilestream/test_tfilestream_string_length_prefixed_serialization
// origin: languages/pascal/tests/pascal/test_pascal_tstream_tfilestream.rs
program Test;
{$mode delphi}
uses Classes;
procedure WriteString(stream: TStream; const s: String);
var len: Integer;
begin
  len := Length(s);
  stream.WriteBuffer(len, SizeOf(Integer));
  if len > 0 then stream.WriteBuffer(s[1], len);
end;
function ReadString(stream: TStream): String;
var len: Integer;
begin
  stream.ReadBuffer(len, SizeOf(Integer));
  SetLength(Result, len);
  if len > 0 then stream.ReadBuffer(Result[1], len);
end;
var fs: TFileStream;
begin
  fs := TFileStream.Create('test_str.dat', fmCreate);
  WriteString(fs, 'StreamedStringData');
  fs.Position := 0;
  WriteLn(ReadString(fs));
  fs.Free;
end.

// vybe-test: pascal/pascal_untyped_file_blockread/test_untyped_file_multidimensional_array_blockwrite
// origin: languages/pascal/tests/pascal/test_pascal_untyped_file_blockread.rs
program Test;
{$mode delphi}
uses SysUtils;
var f: file;
    mat1, mat2: array[0..1, 0..1] of Integer;
    written, readBytes: Integer;
begin
  mat1[0,0] := 1; mat1[0,1] := 2; mat1[1,0] := 3; mat1[1,1] := 4;
  AssignFile(f, 'test_mat_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, mat1, SizeOf(mat1), written);
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, mat2, SizeOf(mat2), readBytes);
  WriteLn(mat2[1,1]);
  CloseFile(f);
end.

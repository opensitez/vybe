// vybe-test: pascal/pascal_file_system_directory_ops/test_filesystem_findfirst_findnext_findclose
// origin: languages/pascal/tests/pascal/test_pascal_file_system_directory_ops.rs
program Test;
{$mode delphi}
uses SysUtils;
var sr: TSearchRec; f: TextFile; count: Integer;
begin
  AssignFile(f, 'test_search_1.tmp'); Rewrite(f); CloseFile(f);
  AssignFile(f, 'test_search_2.tmp'); Rewrite(f); CloseFile(f);

  count := 0;
  if FindFirst('test_search_*.tmp', faAnyFile, sr) = 0 then
  begin
    repeat
      Inc(count);
    until FindNext(sr) <> 0;
    FindClose(sr);
  end;
  WriteLn(count);

  DeleteFile('test_search_1.tmp');
  DeleteFile('test_search_2.tmp');
end.

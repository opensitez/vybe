use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 73: File System Pathing & Directory Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_filesystem_extractfilepath_filename_ext() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var p: String;
begin
  p := '/usr/local/bin/compiler.pas';
  WriteLn(ExtractFileName(p));
  WriteLn(ExtractFileExt(p));
  WriteLn(ExtractFileDir(p));
end.
"#);
    assert_eq!(out, vec!["compiler.pas", ".pas", "/usr/local/bin"]);
}

#[test]
fn test_filesystem_changefileext() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(ChangeFileExt('main.pas', '.exe'));
  WriteLn(ChangeFileExt('data.tar.gz', '.zip'));
end.
"#);
    assert_eq!(out, vec!["main.exe", "data.tar.zip"]);
}

#[test]
fn test_filesystem_fileexists_deletefile() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var f: TextFile; fname: String;
begin
  fname := 'test_fs_del.tmp';
  AssignFile(f, fname); Rewrite(f); WriteLn(f, 'tmp'); CloseFile(f);

  WriteLn(FileExists(fname));
  DeleteFile(fname);
  WriteLn(FileExists(fname));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_filesystem_createdir_directoryexists_removedir() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var dname: String;
begin
  dname := 'test_dir_tmp';
  if DirectoryExists(dname) then RemoveDir(dname);

  CreateDir(dname);
  WriteLn(DirectoryExists(dname));
  RemoveDir(dname);
  WriteLn(DirectoryExists(dname));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_filesystem_forcedirectories() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var path: String;
begin
  path := 'test_parent_dir/test_sub_dir';
  ForceDirectories(path);
  WriteLn(DirectoryExists(path));
  RemoveDir('test_parent_dir/test_sub_dir');
  RemoveDir('test_parent_dir');
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filesystem_getcurrentdir() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(Length(GetCurrentDir) > 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filesystem_renamefile() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var f: TextFile;
begin
  AssignFile(f, 'test_rename_1.tmp'); Rewrite(f); WriteLn(f, 'data'); CloseFile(f);
  RenameFile('test_rename_1.tmp', 'test_rename_2.tmp');
  WriteLn(FileExists('test_rename_1.tmp'));
  WriteLn(FileExists('test_rename_2.tmp'));
  DeleteFile('test_rename_2.tmp');
end.
"#);
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn test_filesystem_expandfilename() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var fullPath: String;
begin
  fullPath := ExpandFileName('relative_file.txt');
  WriteLn(Pos('relative_file.txt', fullPath) > 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filesystem_findfirst_findnext_findclose() {
    let out = run_pascal(r#"
program Test;
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
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_filesystem_include_exclude_trailing_path_delimiter() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var p1, p2: String;
begin
  p1 := IncludeTrailingPathDelimiter('/usr/local');
  p2 := ExcludeTrailingPathDelimiter('/usr/local/');
  WriteLn(p1);
  WriteLn(p2);
end.
"#);
    assert_eq!(out, vec!["/usr/local/", "/usr/local"]);
}

#[test]
fn test_filesystem_ispathdelimiter() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(IsPathDelimiter('/usr/bin', 1));
  WriteLn(IsPathDelimiter('usr/bin', 2));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_filesystem_matchesmask() {
    let out = run_pascal(r#"
program Test;
uses Masks;
begin
  WriteLn(MatchesMask('document.pdf', '*.pdf'));
  WriteLn(MatchesMask('document.txt', '*.pdf'));
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_filesystem_filegetattr_filesetattr() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var f: TextFile; fname: String; attr: Integer;
begin
  fname := 'test_attr.tmp';
  AssignFile(f, fname); Rewrite(f); CloseFile(f);

  attr := FileGetAttr(fname);
  WriteLn(attr <> -1);

  DeleteFile(fname);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filesystem_fileage() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var f: TextFile; fname: String; age: Integer;
begin
  fname := 'test_age.tmp';
  AssignFile(f, fname); Rewrite(f); CloseFile(f);

  age := FileAge(fname);
  WriteLn(age <> -1);

  DeleteFile(fname);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filesystem_samefilename() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(SameFileName('file.txt', 'file.txt'));
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filesystem_extractfiledrive() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  WriteLn(Length(ExtractFileDrive('/usr/bin')) >= 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_filesystem_findclose_safety() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var sr: TSearchRec;
begin
  if FindFirst('non_existent_pattern_xyz_123.*', faAnyFile, sr) <> 0 then
    WriteLn('NotFound');
  FindClose(sr);
end.
"#);
    assert_eq!(out, vec!["NotFound"]);
}

#[test]
fn test_filesystem_tpath_combine() {
    let out = run_pascal(r#"
program Test;
uses System.IOUtils;
begin
  WriteLn(TPath.Combine('/folder', 'file.txt'));
end.
"#);
    assert_eq!(out, vec!["/folder/file.txt"]);
}

#[test]
fn test_filesystem_tfile_fetchalltext() {
    let out = run_pascal(r#"
program Test;
uses System.IOUtils;
var fname: String;
begin
  fname := 'test_tfile.tmp';
  TFile.WriteAllText(fname, 'TFileContent');
  WriteLn(TFile.ReadAllText(fname));
  TFile.Delete(fname);
end.
"#);
    assert_eq!(out, vec!["TFileContent"]);
}

#[test]
fn test_filesystem_tdirectory_exists() {
    let out = run_pascal(r#"
program Test;
uses System.IOUtils;
begin
  WriteLn(TDirectory.Exists('.'));
end.
"#);
    assert_eq!(out, vec!["True"]);
}

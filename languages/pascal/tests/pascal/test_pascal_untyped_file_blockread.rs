use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 63: Untyped Files & BlockRead/BlockWrite Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_untyped_file_blockwrite_blockread_byte_level() {
    let out = run_pascal(r#"
program Test;
var f: file;
    writeBuf, readBuf: array[0..3] of Byte;
    written, readBytes: Integer;
begin
  writeBuf[0] := 10; writeBuf[1] := 20; writeBuf[2] := 30; writeBuf[3] := 40;
  AssignFile(f, 'test_untyped_1.bin');
  Rewrite(f, 1);
  BlockWrite(f, writeBuf[0], 4, written);
  CloseFile(f);
  WriteLn(written);

  Reset(f, 1);
  BlockRead(f, readBuf[0], 4, readBytes);
  CloseFile(f);
  WriteLn(readBytes);
  WriteLn(readBuf[0]);
  WriteLn(readBuf[3]);
end.
"#);
    assert_eq!(out, vec!["4", "4", "10", "40"]);
}

#[test]
fn test_untyped_file_custom_block_size() {
    let out = run_pascal(r#"
program Test;
type TBlock = array[0..15] of Byte;
var f: file;
    blk: TBlock;
    written, readBlocks: Integer;
begin
  FillChar(blk, SizeOf(TBlock), 65);
  AssignFile(f, 'test_block_16.bin');
  Rewrite(f, 16);
  BlockWrite(f, blk, 1, written);
  CloseFile(f);

  Reset(f, 16);
  BlockRead(f, blk, 1, readBlocks);
  CloseFile(f);
  WriteLn(readBlocks);
  WriteLn(blk[0]);
end.
"#);
    assert_eq!(out, vec!["1", "65"]);
}

#[test]
fn test_untyped_file_chunk_copy_loop() {
    let out = run_pascal(r#"
program Test;
var srcFile, dstFile: file;
    buf: array[0..63] of Byte;
    readCount, writtenCount: Integer;
begin
  buf[0] := 99; buf[63] := 88;
  AssignFile(srcFile, 'test_chunk_src.bin');
  Rewrite(srcFile, 1);
  BlockWrite(srcFile, buf[0], 64, writtenCount);
  CloseFile(srcFile);

  Reset(srcFile, 1);
  AssignFile(dstFile, 'test_chunk_dst.bin');
  Rewrite(dstFile, 1);

  repeat
    BlockRead(srcFile, buf[0], 64, readCount);
    if readCount > 0 then
      BlockWrite(dstFile, buf[0], readCount, writtenCount);
  until readCount = 0;

  CloseFile(srcFile); CloseFile(dstFile);
  WriteLn('ChunkCopyComplete');
end.
"#);
    assert_eq!(out, vec!["ChunkCopyComplete"]);
}

#[test]
fn test_untyped_file_filesize_and_filepos_bytes() {
    let out = run_pascal(r#"
program Test;
var f: file;
    buf: array[0..9] of Byte;
    written: Integer;
begin
  AssignFile(f, 'test_sz.bin');
  Rewrite(f, 1);
  BlockWrite(f, buf[0], 10, written);
  CloseFile(f);

  Reset(f, 1);
  WriteLn(FileSize(f));
  WriteLn(FilePos(f));
  BlockRead(f, buf[0], 5, written);
  WriteLn(FilePos(f));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["10", "0", "5"]);
}

#[test]
fn test_untyped_file_seek_block() {
    let out = run_pascal(r#"
program Test;
var f: file;
    buf: array[0..3] of Byte;
    val: Byte; written, readCount: Integer;
begin
  buf[0] := 1; buf[1] := 2; buf[2] := 3; buf[3] := 4;
  AssignFile(f, 'test_seek_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, buf[0], 4, written);
  CloseFile(f);

  Reset(f, 1);
  Seek(f, 2);
  BlockRead(f, val, 1, readCount);
  WriteLn(val);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_untyped_file_record_struct_blockwrite() {
    let out = run_pascal(r#"
program Test;
type THeader = packed record
  Magic: Word;
  Version: Byte;
end;
var f: file;
    h1, h2: THeader;
    written, readBytes: Integer;
begin
  h1.Magic := $4D5A; h1.Version := 2;
  AssignFile(f, 'test_hdr.bin');
  Rewrite(f, 1);
  BlockWrite(f, h1, SizeOf(THeader), written);
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, h2, SizeOf(THeader), readBytes);
  WriteLn(h2.Magic);
  WriteLn(h2.Version);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["19802", "2"]);
}

#[test]
fn test_untyped_file_truncate_at_offset() {
    let out = run_pascal(r#"
program Test;
var f: file;
    buf: array[0..9] of Byte;
    written: Integer;
begin
  AssignFile(f, 'test_trunc_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, buf[0], 10, written);
  CloseFile(f);

  Reset(f, 1);
  Seek(f, 4);
  Truncate(f);
  CloseFile(f);

  Reset(f, 1);
  WriteLn(FileSize(f));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_untyped_file_ioresult_checking() {
    let out = run_pascal(r#"
program Test;
var f: file; err: Integer;
begin
  AssignFile(f, 'non_existent_untyped.bin');
  {$I-}
  Reset(f, 1);
  err := IOResult;
  {$I+}
  WriteLn(err <> 0);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_untyped_file_protection_finally() {
    let out = run_pascal(r#"
program Test;
var f: file; b: Byte; written: Integer;
begin
  b := 42;
  AssignFile(f, 'test_prot_untyped.bin');
  Rewrite(f, 1);
  try
    BlockWrite(f, b, 1, written);
  finally
    CloseFile(f);
    WriteLn('UntypedFileClosedInFinally');
  end;
end.
"#);
    assert_eq!(out, vec!["UntypedFileClosedInFinally"]);
}

#[test]
fn test_untyped_file_default_record_size_128() {
    let out = run_pascal(r#"
program Test;
var f: file;
begin
  AssignFile(f, 'test_def_rec.bin');
  Rewrite(f); // Default 128 bytes
  WriteLn('DefaultRecSize128Initialized');
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["DefaultRecSize128Initialized"]);
}

#[test]
fn test_untyped_file_string_bytes_blockwrite() {
    let out = run_pascal(r#"
program Test;
var f: file;
    text: String;
    readText: String;
    written, readBytes: Integer;
begin
  text := 'RawStringBytes';
  AssignFile(f, 'test_str_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, text[1], Length(text), written);
  CloseFile(f);

  Reset(f, 1);
  SetLength(readText, Length(text));
  BlockRead(f, readText[1], Length(text), readBytes);
  WriteLn(readText);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["RawStringBytes"]);
}

#[test]
fn test_untyped_file_eof_check() {
    let out = run_pascal(r#"
program Test;
var f: file; b: Byte; written, readCount: Integer;
begin
  b := 1;
  AssignFile(f, 'test_eof_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, b, 1, written);
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, b, 1, readCount);
  WriteLn(Eof(f));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_untyped_file_procedure_buffer_parameter() {
    let out = run_pascal(r#"
program Test;
procedure WriteRawData(var f: file; const buffer; size: Integer);
var written: Integer;
begin
  BlockWrite(f, buffer, size, written);
end;
var f: file; val: Integer; readVal: Integer; readBytes: Integer;
begin
  val := 999;
  AssignFile(f, 'test_raw_param.bin');
  Rewrite(f, 1);
  WriteRawData(f, val, SizeOf(Integer));
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, readVal, SizeOf(Integer), readBytes);
  WriteLn(readVal);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["999"]);
}

#[test]
fn test_untyped_file_erase() {
    let out = run_pascal(r#"
program Test;
var f: file; b: Byte; written: Integer;
begin
  b := 55;
  AssignFile(f, 'test_erase_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, b, 1, written);
  CloseFile(f);
  Erase(f);
  WriteLn('ErasedUntypedFile');
end.
"#);
    assert_eq!(out, vec!["ErasedUntypedFile"]);
}

#[test]
fn test_untyped_file_rename() {
    let out = run_pascal(r#"
program Test;
var f: file; b: Byte; written: Integer;
begin
  b := 77;
  AssignFile(f, 'test_orig_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, b, 1, written);
  CloseFile(f);
  Rename(f, 'test_renamed_raw.bin');
  WriteLn('RenamedUntypedFile');
end.
"#);
    assert_eq!(out, vec!["RenamedUntypedFile"]);
}

#[test]
fn test_untyped_file_append_via_seek() {
    let out = run_pascal(r#"
program Test;
var f: file; b: Byte; written, readCount: Integer;
begin
  AssignFile(f, 'test_app_raw.bin');
  Rewrite(f, 1);
  b := 10; BlockWrite(f, b, 1, written);
  CloseFile(f);

  Reset(f, 1);
  Seek(f, FileSize(f));
  b := 20; BlockWrite(f, b, 1, written);
  CloseFile(f);

  Reset(f, 1);
  WriteLn(FileSize(f));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_untyped_file_read_partial_last_block() {
    let out = run_pascal(r#"
program Test;
var f: file;
    buf: array[0..2] of Byte;
    readCount: Integer;
begin
  AssignFile(f, 'test_part.bin');
  Rewrite(f, 1);
  buf[0] := 1; buf[1] := 2;
  BlockWrite(f, buf[0], 2, readCount);
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, buf[0], 5, readCount);
  WriteLn(readCount);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_untyped_file_float_blockwrite() {
    let out = run_pascal(r#"
program Test;
var f: file; r, readR: Real; written, readBytes: Integer;
begin
  r := 98.76;
  AssignFile(f, 'test_real_raw.bin');
  Rewrite(f, 1);
  BlockWrite(f, r, SizeOf(Real), written);
  CloseFile(f);

  Reset(f, 1);
  BlockRead(f, readR, SizeOf(Real), readBytes);
  WriteLn(readR);
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["98.76"]);
}

#[test]
fn test_untyped_file_multidimensional_array_blockwrite() {
    let out = run_pascal(r#"
program Test;
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
"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_untyped_file_empty_file_seek() {
    let out = run_pascal(r#"
program Test;
var f: file;
begin
  AssignFile(f, 'test_empty_seek.bin');
  Rewrite(f, 1);
  CloseFile(f);

  Reset(f, 1);
  Seek(f, 0);
  WriteLn(FilePos(f));
  CloseFile(f);
end.
"#);
    assert_eq!(out, vec!["0"]);
}

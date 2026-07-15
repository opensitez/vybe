//! Core Pascal file I/O lifecycle tests.
use super::helpers::run_pascal;

#[test]
fn textfile_assignfile_rewrite_reset_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin AssignFile(f,'core_assign.txt'); Rewrite(f); WriteLn(f,'alpha'); CloseFile(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["alpha"]
    );
}

#[test]
fn text_alias_rewrite_reset_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: Text; s: string; begin Assign(f,'core_text_alias.txt'); Rewrite(f); WriteLn(f,'alias'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["alias"]
    );
}

#[test]
fn textfile_rewrite_truncates_existing_content() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'core_trunc.txt'); Rewrite(f); WriteLn(f,'old'); Close(f); Rewrite(f); WriteLn(f,'new'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["new"]
    );
}

#[test]
fn textfile_append_preserves_existing_content() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'core_append.txt'); Rewrite(f); WriteLn(f,'first'); Close(f); Append(f); WriteLn(f,'second'); Close(f); Reset(f); ReadLn(f,s); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["second"]
    );
}

#[test]
fn textfile_two_handles_independent_positions_same_name() {
    assert_eq!(
        run_pascal(
            r#"program T; var writer, reader: TextFile; s: string; begin Assign(writer,'core_shared.txt'); Rewrite(writer); WriteLn(writer,'shared'); Close(writer); Assign(reader,'core_shared.txt'); Reset(reader); ReadLn(reader,s); Close(reader); WriteLn(s); end."#
        ),
        &["shared"]
    );
}

#[test]
fn textfile_two_files_keep_separate_contents() {
    assert_eq!(
        run_pascal(
            r#"program T; var a, b: TextFile; s: string; begin Assign(a,'core_a.txt'); Rewrite(a); WriteLn(a,'A'); Close(a); Assign(b,'core_b.txt'); Rewrite(b); WriteLn(b,'B'); Close(b); Reset(a); ReadLn(a,s); Close(a); WriteLn(s); Reset(b); ReadLn(b,s); Close(b); WriteLn(s); end."#
        ),
        &["A", "B"]
    );
}

#[test]
fn textfile_close_is_idempotent_for_reopen() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'core_close.txt'); Rewrite(f); WriteLn(f,'ok'); Close(f); CloseFile(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["ok"]
    );
}

#[test]
fn textfile_var_param_writer_uses_same_handle() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Save(var f: TextFile; msg: string); begin WriteLn(f,msg); end; var f: TextFile; s: string; begin Assign(f,'core_proc.txt'); Rewrite(f); Save(f,'via proc'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["via proc"]
    );
}

#[test]
fn textfile_function_reads_global_handle() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; function Load: string; var s: string; begin Reset(f); ReadLn(f,s); Close(f); Result := s; end; begin Assign(f,'core_func.txt'); Rewrite(f); WriteLn(f,'from func'); Close(f); WriteLn(Load); end."#
        ),
        &["from func"]
    );
}

#[test]
fn textfile_eof_becomes_true_after_last_readln() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'core_eof.txt'); Rewrite(f); WriteLn(f,'last'); Close(f); Reset(f); if not Eof(f) then WriteLn('before'); ReadLn(f,s); if Eof(f) then WriteLn('after'); Close(f); end."#
        ),
        &["before", "after"]
    );
}

#[test]
fn typed_file_integer_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: file of Integer; n: Integer; begin Assign(f,'core_ints.dat'); Rewrite(f); n := 7; Write(f,n); Close(f); Reset(f); n := 0; Read(f,n); Close(f); WriteLn(n); end."#
        ),
        &["7"]
    );
}

#[test]
fn typed_file_alias_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIntFile = file of Integer; var f: TIntFile; n: Integer; begin Assign(f,'core_alias.dat'); Rewrite(f); n := 42; Write(f,n); Close(f); Reset(f); n := 0; Read(f,n); Close(f); WriteLn(n); end."#
        ),
        &["42"]
    );
}

#[test]
fn typed_file_multiple_records_preserve_order() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIntFile = file of Integer; var f: TIntFile; a,b: Integer; begin Assign(f,'core_multi.dat'); Rewrite(f); a := 3; b := 4; Write(f,a); Write(f,b); Close(f); Reset(f); a := 0; b := 0; Read(f,a); Read(f,b); Close(f); WriteLn(a+b); end."#
        ),
        &["7"]
    );
}

#[test]
fn typed_file_eof_after_last_record() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIntFile = file of Integer; var f: TIntFile; n: Integer; begin Assign(f,'core_typed_eof.dat'); Rewrite(f); n := 1; Write(f,n); Close(f); Reset(f); Read(f,n); if Eof(f) then WriteLn('eof'); Close(f); end."#
        ),
        &["eof"]
    );
}

#[test]
fn typed_file_char_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; type TCharFile = file of Char; var f: TCharFile; c: Char; begin Assign(f,'core_chars.dat'); Rewrite(f); c := 'Z'; Write(f,c); Close(f); Reset(f); c := 'X'; Read(f,c); Close(f); WriteLn(c); end."#
        ),
        &["Z"]
    );
}

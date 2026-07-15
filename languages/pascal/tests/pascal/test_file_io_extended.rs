//! Extended Pascal file scenarios that are supported by the current runtime path.
use super::helpers::run_pascal;

#[test]
fn typed_file_byte_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; type TByteFile = file of Byte; var f: TByteFile; b: Byte; begin Assign(f,'ext_byte.dat'); Rewrite(f); b := 255; Write(f,b); Close(f); Reset(f); b := 0; Read(f,b); Close(f); WriteLn(b); end."#
        ),
        &["255"]
    );
}

#[test]
fn typed_file_real_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; type TRealFile = file of Real; var f: TRealFile; r: Real; begin Assign(f,'ext_real.dat'); Rewrite(f); r := 2.5; Write(f,r); Close(f); Reset(f); r := 0; Read(f,r); Close(f); WriteLn(Trunc(r*10)); end."#
        ),
        &["25"]
    );
}

#[test]
fn typed_file_boolean_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBoolFile = file of Boolean; var f: TBoolFile; b: Boolean; begin Assign(f,'ext_bool.dat'); Rewrite(f); b := True; Write(f,b); Close(f); Reset(f); b := False; Read(f,b); Close(f); if b then WriteLn('true'); end."#
        ),
        &["true"]
    );
}

#[test]
fn typed_file_rewrite_truncates_all_records() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIntFile = file of Integer; var f: TIntFile; n: Integer; begin Assign(f,'ext_trunc.dat'); Rewrite(f); n := 1; Write(f,n); n := 2; Write(f,n); Close(f); Rewrite(f); n := 9; Write(f,n); Close(f); Reset(f); n := 0; Read(f,n); Close(f); WriteLn(n); end."#
        ),
        &["9"]
    );
}

#[test]
fn typed_file_write_multiple_arguments_appends_records() {
    assert_eq!(
        run_pascal(
            r#"program T; type TIntFile = file of Integer; var f: TIntFile; a,b,c: Integer; begin Assign(f,'ext_multi_args.dat'); Rewrite(f); a := 1; b := 2; c := 3; Write(f,a,b,c); Close(f); Reset(f); a := 0; b := 0; c := 0; Read(f,a,b,c); Close(f); WriteLn(a+b+c); end."#
        ),
        &["6"]
    );
}

#[test]
fn flush_preserves_data_before_close() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'ext_flush.txt'); Rewrite(f); WriteLn(f,'flushed'); Flush(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["flushed"]
    );
}

#[test]
fn fileexists_reports_rewrite_created_file() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'ext_exists.txt'); Rewrite(f); Close(f); if FileExists('ext_exists.txt') then WriteLn('exists'); end."#
        ),
        &["exists"]
    );
}

#[test]
fn erase_removes_named_file() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'ext_erase.txt'); Rewrite(f); WriteLn(f,'gone'); Close(f); Erase(f); if not FileExists('ext_erase.txt') then WriteLn('gone'); end."#
        ),
        &["gone"]
    );
}

#[test]
fn rename_moves_named_file_contents() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'ext_old.txt'); Rewrite(f); WriteLn(f,'moved'); Close(f); Rename(f,'ext_new.txt'); if not FileExists('ext_old.txt') then WriteLn('old gone'); Assign(f,'ext_new.txt'); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["old gone", "moved"]
    );
}

#[test]
fn reset_after_rename_uses_new_file_name() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'ext_rn1.txt'); Rewrite(f); WriteLn(f,'rn'); Close(f); Rename(f,'ext_rn2.txt'); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["rn"]
    );
}

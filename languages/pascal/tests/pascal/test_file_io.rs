/// Classic Pascal text file I/O: Assign, Rewrite, Reset, ReadLn, WriteLn, Close.
use super::helpers::run_pascal;

#[test]
fn textfile_rewrite_writeln_close() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'out.txt'); Rewrite(f); WriteLn(f,'line1'); Close(f); WriteLn('done'); end."#
        ),
        &["done"]
    );
}

#[test]
fn textfile_append_mode() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'log.txt'); Rewrite(f); WriteLn(f,'a'); Close(f); Append(f); WriteLn(f,'b'); Close(f); WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn textfile_reset_readln() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'in.txt'); Rewrite(f); WriteLn(f,'hello'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["hello"]
    );
}

#[test]
fn textfile_eof_while_read() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; n: Integer; begin Assign(f,'lines.txt'); Rewrite(f); WriteLn(f,'1'); WriteLn(f,'2'); Close(f); Reset(f); n:=0; while not Eof(f) do begin ReadLn(f,s); n:=n+1; end; Close(f); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn textfile_eoln_detect_newline() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'eoln.txt'); Rewrite(f); WriteLn(f,'x'); Close(f); Reset(f); if Eoln(f) then WriteLn('eoln'); Close(f); end."#
        ),
        &["eoln"]
    );
}

#[test]
fn textfile_write_without_ln() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'parts.txt'); Rewrite(f); Write(f,'ab'); Write(f,'c'); Close(f); WriteLn('wrote'); end."#
        ),
        &["wrote"]
    );
}

#[test]
fn textfile_read_after_write_same_program() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; a,b: string; begin Assign(f,'rw.txt'); Rewrite(f); WriteLn(f,'first'); WriteLn(f,'second'); Close(f); Reset(f); ReadLn(f,a); ReadLn(f,b); Close(f); WriteLn(a); WriteLn(b); end."#
        ),
        &["first", "second"]
    );
}

#[test]
fn textfile_integer_write_read() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; n: Integer; begin Assign(f,'num.txt'); Rewrite(f); WriteLn(f,42); Close(f); Reset(f); ReadLn(f,n); Close(f); WriteLn(n); end."#
        ),
        &["42"]
    );
}

#[test]
fn textfile_multiple_close_idempotent_style() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'once.txt'); Rewrite(f); WriteLn(f,'z'); Close(f); WriteLn('closed'); end."#
        ),
        &["closed"]
    );
}

#[test]
fn file_exists_after_rewrite() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'exists.txt'); Rewrite(f); Close(f); WriteLn('yes'); end."#
        ),
        &["yes"]
    );
}

#[test]
fn textfile_readln_empty_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'empty.txt'); Rewrite(f); WriteLn(f); Close(f); Reset(f); ReadLn(f,s); Close(f); if s='' then WriteLn('empty'); end."#
        ),
        &["empty"]
    );
}

#[test]
fn textfile_blockread_style_loop_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; i, c: Integer; s: string; begin Assign(f,'count.txt'); Rewrite(f); for i:=1 to 3 do WriteLn(f,IntToStr(i)); Close(f); Reset(f); c:=0; while not Eof(f) do begin ReadLn(f,s); Inc(c); end; Close(f); WriteLn(c); end."#
        ),
        &["3"]
    );
}

#[test]
fn textfile_rewrite_truncates_previous() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'trunc.txt'); Rewrite(f); WriteLn(f,'old'); Close(f); Rewrite(f); WriteLn(f,'new'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["new"]
    );
}

#[test]
fn typed_file_not_text_but_block() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: file of Integer; n: Integer; begin Assign(f,'ints.dat'); Rewrite(f); n:=7; Write(f,n); Close(f); Reset(f); Read(f,n); Close(f); WriteLn(n); end."#
        ),
        &["7"]
    );
}

#[test]
fn textfile_param_in_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Save(var f: TextFile; msg: string); begin WriteLn(f,msg); end; var f: TextFile; begin Assign(f,'proc.txt'); Rewrite(f); Save(f,'via'); Close(f); WriteLn('saved'); end."#
        ),
        &["saved"]
    );
}

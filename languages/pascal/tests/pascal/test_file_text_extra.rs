//! Focused TextFile scenarios supported by the current Pascal file lowering.
use super::helpers::run_pascal;

#[test]
fn text_readln_reads_whole_line_with_spaces() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'text_line.txt'); Rewrite(f); WriteLn(f,'one two three'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["one two three"]
    );
}

#[test]
fn text_blank_line_roundtrips_as_empty_string() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'text_blank.txt'); Rewrite(f); WriteLn(f); WriteLn(f,'x'); Close(f); Reset(f); ReadLn(f,s); if s = '' then WriteLn('blank'); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["blank", "x"]
    );
}

#[test]
fn text_writeln_multiple_arguments_are_concatenated() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'text_multi_write.txt'); Rewrite(f); WriteLn(f,'A',1,'B',2); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["A1B2"]
    );
}

#[test]
fn text_writeln_no_arguments_writes_empty_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'text_empty_line.txt'); Rewrite(f); WriteLn(f); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Length(s)); end."#
        ),
        &["0"]
    );
}

#[test]
fn text_boolean_token_reads_boolean_target() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; b: Boolean; begin Assign(f,'text_bool.txt'); Rewrite(f); WriteLn(f,'True'); Close(f); Reset(f); ReadLn(f,b); Close(f); if b then WriteLn('true'); end."#
        ),
        &["true"]
    );
}

#[test]
fn text_eof_while_loop_counts_blank_and_nonblank_lines() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; n: Integer; begin Assign(f,'text_count.txt'); Rewrite(f); WriteLn(f); WriteLn(f,'a'); WriteLn(f,'bb'); Close(f); Reset(f); n := 0; while not Eof(f) do begin ReadLn(f,s); n := n + 1; end; Close(f); WriteLn(n); end."#
        ),
        &["3"]
    );
}

#[test]
fn text_long_line_roundtrips_without_truncation() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'text_long.txt'); Rewrite(f); WriteLn(f,'abcdefghijklmnopqrstuvwxyz'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Length(s)); end."#
        ),
        &["26"]
    );
}

#[test]
fn text_filter_loop_reads_each_line_once() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; n: Integer; begin Assign(f,'text_filter.txt'); Rewrite(f); WriteLn(f,'ERR one'); WriteLn(f,'OK two'); WriteLn(f,'ERR three'); Close(f); Reset(f); n := 0; while not Eof(f) do begin ReadLn(f,s); if Copy(s,1,3) = 'ERR' then n := n + 1; end; Close(f); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn text_read_into_array_elements() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; lines: array[0..1] of string; begin Assign(f,'text_array.txt'); Rewrite(f); WriteLn(f,'first'); WriteLn(f,'second'); Close(f); Reset(f); ReadLn(f,lines[0]); ReadLn(f,lines[1]); Close(f); WriteLn(lines[1]); WriteLn(lines[0]); end."#
        ),
        &["second", "first"]
    );
}

#[test]
fn text_nested_reader_procedure_consumes_to_eof() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; procedure ReadAll(var total: Integer); var s: string; begin total := 0; while not Eof(f) do begin ReadLn(f,s); total := total + Length(s); end; end; var n: Integer; begin Assign(f,'text_nested.txt'); Rewrite(f); WriteLn(f,'ab'); WriteLn(f,'cde'); Close(f); Reset(f); ReadAll(n); Close(f); WriteLn(n); end."#
        ),
        &["5"]
    );
}

#[test]
fn text_readln_without_variable_discards_one_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'text_discard.txt'); Rewrite(f); WriteLn(f,'skip'); WriteLn(f,'keep'); Close(f); Reset(f); ReadLn(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["keep"]
    );
}

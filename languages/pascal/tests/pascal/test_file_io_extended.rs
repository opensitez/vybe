/// Extended TextFile and typed file I/O scenarios.
use super::helpers::run_pascal;

#[test]
fn textfile_write_multiple_lines_then_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; n: Integer; begin Assign(f,'multi.txt'); Rewrite(f); WriteLn(f,'a'); WriteLn(f,'b'); WriteLn(f,'c'); Close(f); Reset(f); n:=0; while not Eof(f) do begin ReadLn(f,s); n:=n+1; end; Close(f); WriteLn(n); end."#
        ),
        &["3"]
    );
}

#[test]
fn textfile_read_first_line_only() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'first.txt'); Rewrite(f); WriteLn(f,'one'); WriteLn(f,'two'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["one"]
    );
}

#[test]
fn textfile_append_preserves_then_adds() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'app.txt'); Rewrite(f); WriteLn(f,'old'); Close(f); Append(f); WriteLn(f,'new'); Close(f); Reset(f); ReadLn(f,s); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["new"]
    );
}

#[test]
fn textfile_rewrite_overwrites_content() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'ovr.txt'); Rewrite(f); WriteLn(f,'v1'); Close(f); Rewrite(f); WriteLn(f,'v2'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["v2"]
    );
}

#[test]
fn textfile_write_integers_separate_lines() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; a, b: Integer; begin Assign(f,'nums.txt'); Rewrite(f); WriteLn(f,10); WriteLn(f,20); Close(f); Reset(f); ReadLn(f,a); ReadLn(f,b); Close(f); WriteLn(a+b); end."#
        ),
        &["30"]
    );
}

#[test]
fn textfile_write_real_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; r: Real; begin Assign(f,'real.txt'); Rewrite(f); WriteLn(f,3.5); Close(f); Reset(f); ReadLn(f,r); Close(f); WriteLn(r > 3); end."#
        ),
        &["True"]
    );
}

#[test]
fn textfile_empty_file_eof_immediate() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'empty.txt'); Rewrite(f); Close(f); Reset(f); if Eof(f) then WriteLn('eof'); Close(f); end."#
        ),
        &["eof"]
    );
}

#[test]
fn textfile_readln_skips_blank_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'blank.txt'); Rewrite(f); WriteLn(f,''); WriteLn(f,'x'); Close(f); Reset(f); ReadLn(f,s); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["x"]
    );
}

#[test]
fn textfile_write_without_ln_then_read() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'parts2.txt'); Rewrite(f); Write(f,'ab'); Write(f,'c'); WriteLn(f,''); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["abc"]
    );
}

#[test]
fn textfile_procedure_writes_file() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; procedure Save(msg: string); begin Assign(f,'proc.txt'); Rewrite(f); WriteLn(f,msg); Close(f); end; var s: string; begin Save('saved'); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["saved"]
    );
}

#[test]
fn textfile_function_reads_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; function Load: string; var s: string; begin Reset(f); ReadLn(f,s); Close(f); Result := s; end; begin Assign(f,'load.txt'); Rewrite(f); WriteLn(f,'data'); Close(f); WriteLn(Load); end."#
        ),
        &["data"]
    );
}

#[test]
fn textfile_two_files_independent() {
    assert_eq!(
        run_pascal(
            r#"program T; var f1, f2: TextFile; s: string; begin Assign(f1,'a.txt'); Rewrite(f1); WriteLn(f1,'A'); Close(f1); Assign(f2,'b.txt'); Rewrite(f2); WriteLn(f2,'B'); Close(f2); Reset(f1); ReadLn(f1,s); Close(f1); WriteLn(s); Reset(f2); ReadLn(f2,s); Close(f2); WriteLn(s); end."#
        ),
        &["A", "B"]
    );
}

#[test]
fn textfile_loop_write_then_sum_read() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; i, n, s: Integer; begin Assign(f,'sum.txt'); Rewrite(f); for i := 1 to 4 do WriteLn(f,i); Close(f); Reset(f); s := 0; while not Eof(f) do begin ReadLn(f,n); s := s + n; end; Close(f); WriteLn(s); end."#
        ),
        &["10"]
    );
}

#[test]
fn textfile_char_by_write_sequence() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'chars.txt'); Rewrite(f); Write(f,'x'); Write(f,'y'); Write(f,'z'); Close(f); WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn textfile_boolean_as_text() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'bool.txt'); Rewrite(f); WriteLn(f,'True'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["True"]
    );
}

#[test]
fn textfile_long_line_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'long.txt'); Rewrite(f); WriteLn(f,'abcdefghijklmnopqrstuvwxyz'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Length(s)); end."#
        ),
        &["26"]
    );
}

#[test]
fn textfile_mixed_write_writeln() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'mix.txt'); Rewrite(f); Write(f,'pre'); WriteLn(f,'post'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["prepost"]
    );
}

#[test]
fn textfile_read_in_for_loop_fixed_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; i: Integer; s, last: string; begin Assign(f,'for.txt'); Rewrite(f); WriteLn(f,'1'); WriteLn(f,'2'); WriteLn(f,'3'); Close(f); Reset(f); for i := 1 to 3 do ReadLn(f,s); last := s; Close(f); WriteLn(last); end."#
        ),
        &["3"]
    );
}

#[test]
fn textfile_eoln_on_partial_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; begin Assign(f,'partial.txt'); Rewrite(f); Write(f,'noeol'); Close(f); Reset(f); if Eof(f) then WriteLn('eof') else WriteLn('data'); Close(f); end."#
        ),
        &["data"]
    );
}

#[test]
fn textfile_assign_different_names() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'name1.txt'); Rewrite(f); WriteLn(f,'n1'); Close(f); Assign(f,'name2.txt'); Rewrite(f); WriteLn(f,'n2'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["n2"]
    );
}

#[test]
fn textfile_csv_style_parse() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; a, b: string; begin Assign(f,'csv.txt'); Rewrite(f); WriteLn(f,'10,20'); Close(f); Reset(f); ReadLn(f,a); Close(f); b := Copy(a, Pos(',', a) + 1, 10); WriteLn(b); end."#
        ),
        &["20"]
    );
}

#[test]
fn textfile_log_timestamp_style() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; n: Integer; begin Assign(f,'log2.txt'); Rewrite(f); WriteLn(f,'[INFO] start'); WriteLn(f,'[INFO] end'); Close(f); Reset(f); n:=0; while not Eof(f) do begin ReadLn(f,s); n:=n+1; end; Close(f); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn textfile_config_key_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; line, key, val: string; p: Integer; begin Assign(f,'cfg.txt'); Rewrite(f); WriteLn(f,'port=8080'); Close(f); Reset(f); ReadLn(f,line); Close(f); p := Pos('=', line); key := Copy(line, 1, p - 1); val := Copy(line, p + 1, 10); WriteLn(key); WriteLn(val); end."#
        ),
        &["port", "8080"]
    );
}

#[test]
fn typed_file_block_integer_array() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF = file of Integer; var f: TF; n: Integer; begin Assign(f,'block.dat'); Rewrite(f); n := 99; Write(f, n); Close(f); Reset(f); Read(f, n); Close(f); WriteLn(n); end."#
        ),
        &["99"]
    );
}

#[test]
fn typed_file_multiple_records() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF = file of Integer; var f: TF; a, b, c: Integer; begin Assign(f,'multi.dat'); Rewrite(f); a := 1; b := 2; Write(f,a); Write(f,b); Close(f); Reset(f); Read(f,a); Read(f,b); Close(f); WriteLn(a+b); end."#
        ),
        &["3"]
    );
}

#[test]
fn typed_file_char_sequence() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF = file of Char; var f: TF; c: Char; begin Assign(f,'chars.dat'); Rewrite(f); c := 'Z'; Write(f,c); Close(f); Reset(f); Read(f,c); Close(f); WriteLn(c); end."#
        ),
        &["Z"]
    );
}

#[test]
fn typed_file_eof_after_last_read() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF = file of Integer; var f: TF; n: Integer; begin Assign(f,'one.dat'); Rewrite(f); n := 7; Write(f,n); Close(f); Reset(f); Read(f,n); if Eof(f) then WriteLn('eof'); Close(f); end."#
        ),
        &["eof"]
    );
}

#[test]
fn typed_file_rewrite_truncates() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF = file of Integer; var f: TF; n: Integer; begin Assign(f,'trunc.dat'); Rewrite(f); n := 1; Write(f,n); Close(f); Rewrite(f); n := 2; Write(f,n); Close(f); Reset(f); Read(f,n); Close(f); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn textfile_copy_line_to_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var src, dst: TextFile; s: string; begin Assign(src,'src.txt'); Rewrite(src); WriteLn(src,'copyme'); Close(src); Assign(dst,'dst.txt'); Rewrite(dst); Reset(src); ReadLn(src,s); WriteLn(dst,s); Close(src); Close(dst); Reset(dst); ReadLn(dst,s); Close(dst); WriteLn(s); end."#
        ),
        &["copyme"]
    );
}

#[test]
fn textfile_filter_lines_by_prefix() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; n: Integer; begin Assign(f,'filt.txt'); Rewrite(f); WriteLn(f,'ERR one'); WriteLn(f,'OK two'); WriteLn(f,'ERR three'); Close(f); Reset(f); n := 0; while not Eof(f) do begin ReadLn(f,s); if Copy(s,1,3) = 'ERR' then n := n + 1; end; Close(f); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn textfile_write_tab_separated() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'tsv.txt'); Rewrite(f); Write(f,'a'); Write(f,#9); WriteLn(f,'b'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Pos(#9, s) > 0); end."#
        ),
        &["True"]
    );
}

#[test]
fn textfile_read_accumulate_words() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; words: Integer; begin Assign(f,'words.txt'); Rewrite(f); WriteLn(f,'one two'); WriteLn(f,'three'); Close(f); Reset(f); words := 0; while not Eof(f) do begin ReadLn(f,s); if s <> '' then words := words + 1; end; Close(f); WriteLn(words); end."#
        ),
        &["2"]
    );
}

#[test]
fn textfile_nested_procedure_reader() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; procedure ReadAll(var total: Integer); var s: string; begin total := 0; while not Eof(f) do begin ReadLn(f,s); total := total + Length(s); end; end; var n: Integer; begin Assign(f,'len.txt'); Rewrite(f); WriteLn(f,'ab'); WriteLn(f,'cde'); Close(f); Reset(f); ReadAll(n); Close(f); WriteLn(n); end."#
        ),
        &["5"]
    );
}

#[test]
fn textfile_write_quoted_string() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; begin Assign(f,'q.txt'); Rewrite(f); WriteLn(f,'"hello"'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["\"hello\""]
    );
}

#[test]
fn typed_file_seek_style_skip_first() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF = file of Integer; var f: TF; a, b: Integer; begin Assign(f,'skip.dat'); Rewrite(f); a := 10; b := 20; Write(f,a); Write(f,b); Close(f); Reset(f); Read(f,a); Read(f,b); Close(f); WriteLn(b); end."#
        ),
        &["20"]
    );
}

#[test]
fn textfile_max_line_length_track() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; s: string; maxLen, len: Integer; begin Assign(f,'maxl.txt'); Rewrite(f); WriteLn(f,'aa'); WriteLn(f,'bbbb'); Close(f); Reset(f); maxLen := 0; while not Eof(f) do begin ReadLn(f,s); len := Length(s); if len > maxLen then maxLen := len; end; Close(f); WriteLn(maxLen); end."#
        ),
        &["4"]
    );
}

#[test]
fn textfile_write_enum_as_line() {
    assert_eq!(
        run_pascal(
            r#"program T; type TS = (On, Off); var f: TextFile; state: TS; s: string; begin state := On; Assign(f,'state.txt'); Rewrite(f); WriteLn(f,Ord(state)); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["0"]
    );
}

#[test]
fn textfile_reverse_lines_on_read() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; lines: array[0..1] of string; begin Assign(f,'rev.txt'); Rewrite(f); WriteLn(f,'first'); WriteLn(f,'second'); Close(f); Reset(f); ReadLn(f,lines[0]); ReadLn(f,lines[1]); Close(f); WriteLn(lines[1]); WriteLn(lines[0]); end."#
        ),
        &["second", "first"]
    );
}

#[test]
fn typed_file_boolean_as_ordinal() {
    assert_eq!(
        run_pascal(
            r#"program T; type TF = file of Byte; var f: TF; b: Byte; begin Assign(f,'bool.dat'); Rewrite(f); b := 1; Write(f,b); Close(f); Reset(f); Read(f,b); Close(f); WriteLn(b); end."#
        ),
        &["1"]
    );
}

#[test]
fn textfile_batch_numbers_min_max() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: TextFile; n, vmin, vmax: Integer; begin Assign(f,'batch.txt'); Rewrite(f); WriteLn(f,5); WriteLn(f,2); WriteLn(f,8); Close(f); Reset(f); vmin := 9999; vmax := -9999; while not Eof(f) do begin ReadLn(f,n); if n < vmin then vmin := n; if n > vmax then vmax := n; end; Close(f); WriteLn(vmin); WriteLn(vmax); end."#
        ),
        &["2", "8"]
    );
}

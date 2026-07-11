/// Additional text file read/write patterns.
use super::helpers::run_pascal;

#[test]
fn text_rewrite_single_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; begin Assign(f,'a1.txt'); Rewrite(f); WriteLn(f,'one'); Close(f); WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn text_append_second_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; begin Assign(f,'a2.txt'); Rewrite(f); WriteLn(f,'a'); Close(f); Append(f); WriteLn(f,'b'); Close(f); WriteLn('done'); end."#
        ),
        &["done"]
    );
}

#[test]
fn text_read_back_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a3.txt'); Rewrite(f); WriteLn(f,'hello'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["hello"]
    );
}

#[test]
fn text_read_two_lines() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; a,b:string; begin Assign(f,'a4.txt'); Rewrite(f); WriteLn(f,'first'); WriteLn(f,'second'); Close(f); Reset(f); ReadLn(f,a); ReadLn(f,b); Close(f); WriteLn(a); WriteLn(b); end."#
        ),
        &["first", "second"]
    );
}

#[test]
fn text_eof_count_lines() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; n:Integer; begin Assign(f,'a5.txt'); Rewrite(f); WriteLn(f,'1'); WriteLn(f,'2'); WriteLn(f,'3'); Close(f); Reset(f); n:=0; while not Eof(f) do begin ReadLn(f,s); Inc(n); end; Close(f); WriteLn(n); end."#
        ),
        &["3"]
    );
}

#[test]
fn text_eoln_detect() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; begin Assign(f,'a6.txt'); Rewrite(f); WriteLn(f,'x'); Close(f); Reset(f); if Eoln(f) then WriteLn('eoln'); Close(f); end."#
        ),
        &["eoln"]
    );
}

#[test]
fn text_write_without_ln() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; begin Assign(f,'a7.txt'); Rewrite(f); Write(f,'ab'); Write(f,'c'); Close(f); WriteLn('w'); end."#
        ),
        &["w"]
    );
}

#[test]
fn text_integer_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; n:Integer; begin Assign(f,'a8.txt'); Rewrite(f); WriteLn(f,42); Close(f); Reset(f); ReadLn(f,n); Close(f); WriteLn(n); end."#
        ),
        &["42"]
    );
}

#[test]
fn text_real_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; r:Real; begin Assign(f,'a9.txt'); Rewrite(f); WriteLn(f,3.5); Close(f); Reset(f); ReadLn(f,r); Close(f); WriteLn(Trunc(r*10)); end."#
        ),
        &["35"]
    );
}

#[test]
fn text_multiple_integers() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; a,b:Integer; begin Assign(f,'a10.txt'); Rewrite(f); WriteLn(f,10); WriteLn(f,20); Close(f); Reset(f); ReadLn(f,a); ReadLn(f,b); Close(f); WriteLn(a+b); end."#
        ),
        &["30"]
    );
}

#[test]
fn text_empty_file_eof() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; begin Assign(f,'a11.txt'); Rewrite(f); Close(f); Reset(f); if Eof(f) then WriteLn('empty'); Close(f); end."#
        ),
        &["empty"]
    );
}

#[test]
fn text_readln_skips_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a12.txt'); Rewrite(f); WriteLn(f,'skip'); WriteLn(f,'keep'); Close(f); Reset(f); ReadLn(f,s); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["keep"]
    );
}

#[test]
fn text_rewrite_truncates() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a13.txt'); Rewrite(f); WriteLn(f,'old'); Close(f); Rewrite(f); WriteLn(f,'new'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["new"]
    );
}

#[test]
fn text_append_preserves() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a14.txt'); Rewrite(f); WriteLn(f,'1'); Close(f); Append(f); WriteLn(f,'2'); Close(f); Reset(f); ReadLn(f,s); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["2"]
    );
}

#[test]
fn text_char_write() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a15.txt'); Rewrite(f); Write(f,'Z'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["Z"]
    );
}

#[test]
fn text_while_read_all() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; t:string; begin Assign(f,'a16.txt'); Rewrite(f); WriteLn(f,'aa'); WriteLn(f,'bb'); Close(f); Reset(f); t:=''; while not Eof(f) do begin ReadLn(f,s); t:=t+s; end; Close(f); WriteLn(t); end."#
        ),
        &["aabb"]
    );
}

#[test]
fn text_three_line_sum_len() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; n:Integer; begin Assign(f,'a17.txt'); Rewrite(f); WriteLn(f,'a'); WriteLn(f,'bb'); WriteLn(f,'ccc'); Close(f); Reset(f); n:=0; while not Eof(f) do begin ReadLn(f,s); n:=n+Length(s); end; Close(f); WriteLn(n); end."#
        ),
        &["6"]
    );
}

#[test]
fn text_write_number_sequence() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; i:Integer; begin Assign(f,'a18.txt'); Rewrite(f); for i:=1 to 3 do WriteLn(f,i); Close(f); WriteLn('seq'); end."#
        ),
        &["seq"]
    );
}

#[test]
fn text_read_first_only() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a19.txt'); Rewrite(f); WriteLn(f,'only'); WriteLn(f,'more'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["only"]
    );
}

#[test]
fn text_mixed_write_read() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; n:Integer; begin Assign(f,'a20.txt'); Rewrite(f); WriteLn(f,7); Close(f); Reset(f); ReadLn(f,n); Close(f); WriteLn(n*2); end."#
        ),
        &["14"]
    );
}

#[test]
fn text_blank_line_count() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; n:Integer; begin Assign(f,'a21.txt'); Rewrite(f); WriteLn(f); WriteLn(f,'x'); Close(f); Reset(f); n:=0; while not Eof(f) do begin ReadLn(f,s); Inc(n); end; Close(f); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn text_write_tabs() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a22.txt'); Rewrite(f); WriteLn(f,'a'+#9+'b'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Length(s)); end."#
        ),
        &["3"]
    );
}

#[test]
fn text_close_before_reopen() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a23.txt'); Rewrite(f); WriteLn(f,'z'); Close(f); Assign(f,'a23.txt'); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["z"]
    );
}

#[test]
fn text_two_assign_same() {
    assert_eq!(
        run_pascal(
            r#"program T; var f,g:TextFile; begin Assign(f,'a24.txt'); Rewrite(f); WriteLn(f,'m'); Close(f); Assign(g,'a24.txt'); Reset(g); WriteLn('r'); Close(g); end."#
        ),
        &["r"]
    );
}

#[test]
fn text_read_integer_list() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; a,b,c:Integer; begin Assign(f,'a25.txt'); Rewrite(f); WriteLn(f,1); WriteLn(f,2); WriteLn(f,3); Close(f); Reset(f); ReadLn(f,a); ReadLn(f,b); ReadLn(f,c); Close(f); WriteLn(a+b+c); end."#
        ),
        &["6"]
    );
}

#[test]
fn text_write_bool_as_text() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a26.txt'); Rewrite(f); WriteLn(f,'true'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["true"]
    );
}

#[test]
fn text_partial_write_flush() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; begin Assign(f,'a27.txt'); Rewrite(f); Write(f,'pre'); WriteLn(f,'fix'); Close(f); WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn text_read_after_append() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a28.txt'); Rewrite(f); WriteLn(f,'base'); Close(f); Append(f); WriteLn(f,'ext'); Close(f); Reset(f); ReadLn(f,s); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["ext"]
    );
}

#[test]
fn text_count_words_lines() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; n:Integer; begin Assign(f,'a29.txt'); Rewrite(f); WriteLn(f,'one two'); WriteLn(f,'three'); Close(f); Reset(f); n:=0; while not Eof(f) do begin ReadLn(f,s); Inc(n); end; Close(f); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn text_rewrite_empty_then_write() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a30.txt'); Rewrite(f); Close(f); Rewrite(f); WriteLn(f,'after'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["after"]
    );
}

#[test]
fn text_long_line_roundtrip() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a31.txt'); Rewrite(f); WriteLn(f,'abcdefghij'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Length(s)); end."#
        ),
        &["10"]
    );
}

#[test]
fn text_sequential_numbers() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; i,n:Integer; begin Assign(f,'a32.txt'); Rewrite(f); for i:=1 to 5 do WriteLn(f,i); Close(f); Reset(f); n:=0; ReadLn(f,n); Close(f); WriteLn(n); end."#
        ),
        &["1"]
    );
}

#[test]
fn text_write_multiple_strings() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a33.txt'); Rewrite(f); WriteLn(f,'p'); WriteLn(f,'q'); WriteLn(f,'r'); Close(f); Reset(f); ReadLn(f,s); ReadLn(f,s); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["r"]
    );
}

#[test]
fn text_readln_trim_len() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a34.txt'); Rewrite(f); WriteLn(f,'abc'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Length(s)); end."#
        ),
        &["3"]
    );
}

#[test]
fn text_append_count_after() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; n:Integer; begin Assign(f,'a35.txt'); Rewrite(f); WriteLn(f,'1'); Close(f); Append(f); WriteLn(f,'2'); Close(f); Reset(f); n:=0; while not Eof(f) do begin ReadLn(f,s); Inc(n); end; Close(f); WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn text_write_int_to_str() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a36.txt'); Rewrite(f); WriteLn(f,IntToStr(88)); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["88"]
    );
}

#[test]
fn text_nested_assign_reset() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a37.txt'); Rewrite(f); WriteLn(f,'data'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["data"]
    );
}

#[test]
fn text_four_line_read_last() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a38.txt'); Rewrite(f); WriteLn(f,'a'); WriteLn(f,'b'); WriteLn(f,'c'); WriteLn(f,'d'); Close(f); Reset(f); ReadLn(f,s); ReadLn(f,s); ReadLn(f,s); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["d"]
    );
}

#[test]
fn text_write_real_format() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a39.txt'); Rewrite(f); WriteLn(f,FormatFloat('0.0',2.5)); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["2.5"]
    );
}

#[test]
fn text_read_empty_line() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a40.txt'); Rewrite(f); WriteLn(f); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(Length(s)); end."#
        ),
        &["0"]
    );
}

#[test]
fn text_multiple_rewrite_cycles() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; s:string; begin Assign(f,'a41.txt'); Rewrite(f); WriteLn(f,'v1'); Close(f); Rewrite(f); WriteLn(f,'v2'); Close(f); Reset(f); ReadLn(f,s); Close(f); WriteLn(s); end."#
        ),
        &["v2"]
    );
}

#[test]
fn text_sum_file_integers() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:TextFile; i,s,n:Integer; begin Assign(f,'a42.txt'); Rewrite(f); for i:=1 to 4 do WriteLn(f,i); Close(f); Reset(f); s:=0; while not Eof(f) do begin ReadLn(f,n); s:=s+n; end; Close(f); WriteLn(s); end."#
        ),
        &["10"]
    );
}

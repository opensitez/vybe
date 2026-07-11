//! bufio: Scanner split modes, custom SplitFunc, Reader ReadString/ReadBytes,
//! Peek, UnreadByte/Rune — extended coverage distinct from `test_bufio_io.rs`.

go_run_cases! {
    scanner_scan_lines_first_row => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a\\nb\\n\")); sc.Split(bufio.ScanLines); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["a"]
    ),
    scanner_scan_lines_second_row => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a\\nb\\n\")); sc.Split(bufio.ScanLines); sc.Scan(); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["b"]
    ),
    scanner_scan_lines_trailing_line_no_newline => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"solo\")); sc.Split(bufio.ScanLines); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["solo"]
    ),
    scanner_scan_lines_empty_line_between => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a\\n\\nb\")); sc.Split(bufio.ScanLines); sc.Scan(); sc.Scan(); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["b"]
    ),
    scanner_scan_lines_count_two => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"x\\ny\\n\")); sc.Split(bufio.ScanLines); n := 0; for sc.Scan() { n++ }; fmt.Println(n) }",
        vec!["2"]
    ),
    scanner_scan_bytes_second_byte => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"AB\")); sc.Split(bufio.ScanBytes); sc.Scan(); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["B"]
    ),
    scanner_scan_bytes_count_three => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"xyz\")); sc.Split(bufio.ScanBytes); n := 0; for sc.Scan() { n++ }; fmt.Println(n) }",
        vec!["3"]
    ),
    scanner_scan_words_skips_whitespace => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"  go  lang  \")); sc.Split(bufio.ScanWords); sc.Scan(); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["lang"]
    ),
    scanner_scan_words_tab_delimited => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a\\tb\")); sc.Split(bufio.ScanWords); sc.Scan(); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["b"]
    ),
    scanner_default_split_first_line => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"line1\\nline2\")); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["line1"]
    ),
    scanner_err_nil_after_success => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"ok\")); sc.Scan(); fmt.Println(sc.Err() == nil) }",
        vec!["true"]
    ),
    scanner_bytes_matches_scan_text => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"tok\")); sc.Scan(); fmt.Println(string(sc.Bytes()) == sc.Text()) }",
        vec!["true"]
    ),
    scanner_switch_split_mid_stream => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a b\")); sc.Split(bufio.ScanWords); sc.Scan(); sc.Split(bufio.ScanBytes); sc.Scan(); fmt.Println(sc.Text()) }",
        vec![" "]
    ),
    reader_read_string_until_tab => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"key\\tval\")); s, _ := r.ReadString('\\t'); fmt.Println(s) }",
        vec!["key\t"]
    ),
    reader_read_string_until_pipe => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"a|b\")); s, _ := r.ReadString('|'); fmt.Println(s) }",
        vec!["a|"]
    ),
    reader_read_string_eof_without_delim => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"tail\")); s, _ := r.ReadString('\\n'); fmt.Println(s) }",
        vec!["tail"]
    ),
    reader_read_bytes_until_comma => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"one,two\")); b, _ := r.ReadBytes(','); fmt.Println(string(b)) }",
        vec!["one,"]
    ),
    reader_read_bytes_until_newline => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"row\\n\")); b, _ := r.ReadBytes('\\n'); fmt.Println(string(b)) }",
        vec!["row\n"]
    ),
    reader_read_bytes_single_byte_delim => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"x:y\")); b, _ := r.ReadBytes(':'); fmt.Println(string(b)) }",
        vec!["x:"]
    ),
    reader_peek_one_byte => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"peek\")); p, _ := r.Peek(1); fmt.Println(string(p)) }",
        vec!["p"]
    ),
    reader_peek_three_bytes => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"peek\")); p, _ := r.Peek(3); fmt.Println(string(p)) }",
        vec!["pee"]
    ),
    reader_peek_does_not_consume => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"data\")); _, _ = r.Peek(2); s, _ := r.ReadString('a'); fmt.Println(s) }",
        vec!["da"]
    ),
    reader_peek_after_partial_read => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"abcd\")); r.ReadByte(); p, _ := r.Peek(2); fmt.Println(string(p)) }",
        vec!["bc"]
    ),
    reader_unread_byte_then_read_rune => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"go\")); b, _ := r.ReadByte(); r.UnreadByte(); ch, _, _ := r.ReadRune(); fmt.Println(string(b), string(ch)) }",
        vec!["g g"]
    ),
    reader_unread_rune_then_read_byte => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"日\")); ch, _, _ := r.ReadRune(); r.UnreadRune(); b, _ := r.ReadByte(); fmt.Println(string(ch), int(b)) }",
        vec!["日 230"]
    ),
    reader_unread_rune_ascii => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"z\")); ch, _, _ := r.ReadRune(); r.UnreadRune(); ch2, _, _ := r.ReadRune(); fmt.Println(string(ch), string(ch2)) }",
        vec!["z z"]
    ),
    reader_double_unread_byte => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"ab\")); b1, _ := r.ReadByte(); r.UnreadByte(); b2, _ := r.ReadByte(); fmt.Println(string(b1), string(b2)) }",
        vec!["a a"]
    ),
    reader_read_string_then_read_byte => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"hi\\nthere\")); _, _ = r.ReadString('\\n'); b, _ := r.ReadByte(); fmt.Println(string(b)) }",
        vec!["t"]
    ),
    reader_read_bytes_then_read_string => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"a,b\")); _, _ = r.ReadBytes(','); s, _ := r.ReadString('\\x00'); fmt.Println(s) }",
        vec!["b"]
    ),
    scanner_words_on_punctuation_stream => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a,b c\")); sc.Split(bufio.ScanWords); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["a,b"]
    ),
    scanner_lines_preserves_inner_spaces => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a b\\nc d\")); sc.Split(bufio.ScanLines); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["a b"]
    ),
    scanner_bytes_on_utf8_rune => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"日\")); sc.Split(bufio.ScanBytes); n := 0; for sc.Scan() { n++ }; fmt.Println(n) }",
        vec!["3"]
    ),
    reader_peek_beyond_buffer_returns_available => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"ab\")); p, _ := r.Peek(10); fmt.Println(len(p), string(p)) }",
        vec!["2 ab"]
    ),
    reader_read_string_empty_prefix => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\",rest\")); s, _ := r.ReadString(','); fmt.Println(len(s), string(s)) }",
        vec!["1 ,"]
    ),
    reader_read_bytes_empty_suffix => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"only\")); b, _ := r.ReadBytes('z'); fmt.Println(string(b)) }",
        vec!["only"]
    ),
    scanner_scan_false_on_blank_input => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"\")); sc.Split(bufio.ScanWords); fmt.Println(sc.Scan()) }",
        vec!["false"]
    ),
    scanner_scan_lines_blank_only => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"\\n\")); sc.Split(bufio.ScanLines); sc.Scan(); fmt.Println(len(sc.Text())) }",
        vec!["0"]
    ),
    reader_unread_byte_after_peek => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"xy\")); p, _ := r.Peek(1); b, _ := r.ReadByte(); fmt.Println(string(p), string(b)) }",
        vec!["x x"]
    ),
    reader_read_rune_size_reported => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"日\")); _, sz, _ := r.ReadRune(); fmt.Println(sz) }",
        vec!["3"]
    ),
    scanner_default_third_line => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"1\\n2\\n3\")); sc.Scan(); sc.Scan(); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["3"]
    ),
    reader_read_string_multibyte_before_delim => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"日;next\")); s, _ := r.ReadString(';'); fmt.Println(s) }",
        vec!["日;"]
    ),
    reader_peek_zero_length => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"z\")); p, _ := r.Peek(0); fmt.Println(len(p)) }",
        vec!["0"]
    ),
    scanner_words_single_token_stream => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"solo\")); sc.Split(bufio.ScanWords); sc.Scan(); fmt.Println(sc.Text()); fmt.Println(sc.Scan()) }",
        vec!["solo", "false"]
    ),
}

go_compile_cases! {
    scanner_custom_split_on_semicolon => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a;b\")); sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { for i, b := range data { if b == ';' { return i + 1, data[:i], nil } }; if atEOF && len(data) > 0 { return len(data), data, nil }; return 0, nil, nil }); _ = sc.Scan() }",
    scanner_custom_split_fixed_width_two => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"abcdef\")); sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { if len(data) >= 2 { return 2, data[:2], nil }; return 0, data, nil }); _ = sc.Scan() }",
    scanner_custom_split_comma_with_remainder => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"x,y,z\")); sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { for i, b := range data { if b == ',' { return i + 1, data[:i], nil } }; return 0, nil, nil }); for sc.Scan() { _ = sc.Text() } }",
    scanner_buffer_reuse => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"tok\")); _ = sc.Buffer(make([]byte, 0, 64), 1024); _ = sc.Scan() }",
    scanner_split_reassign_after_scan => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a b\")); sc.Split(bufio.ScanWords); _ = sc.Scan(); sc.Split(bufio.ScanLines); _ = sc.Scan() }",
    reader_unread_rune_after_read_bytes => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"ab\")); _, _ = r.ReadBytes('a'); _ = r.UnreadRune() }",
    reader_peek_after_read_string => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"a\\nb\")); _, _ = r.ReadString('\\n'); _, _ = r.Peek(1) }",
    reader_read_bytes_then_unread_byte => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"xy\")); _, _ = r.ReadBytes('x'); _ = r.UnreadByte() }",
    scanner_scan_runes_second_rune => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"ab\")); sc.Split(bufio.ScanRunes); _ = sc.Scan(); _ = sc.Scan() }",
    scanner_custom_split_returns_empty_token => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\",a\")); sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { if len(data) > 0 && data[0] == ',' { return 1, []byte{}, nil }; return 0, nil, nil }); _ = sc.Scan() }",
    reader_read_string_delim_null => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"a\\x00b\")); _, _ = r.ReadString('\\x00') }",
    reader_peek_with_small_reader_buffer => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReaderSize(strings.NewReader(\"abcd\"), 2); _, _ = r.Peek(3) }",
    scanner_lines_carriage_return => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a\\r\\nb\")); sc.Split(bufio.ScanLines); for sc.Scan() { _ = sc.Text() } }",
    reader_unread_byte_after_read_rune => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"go\")); _, _, _ = r.ReadRune(); _ = r.UnreadByte() }",
    scanner_custom_split_at_eof_flush => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"end\")); sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { if atEOF { return len(data), data, nil }; return 0, nil, nil }); _ = sc.Scan() }",
    reader_read_bytes_large_delim => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"startEND\")); _, _ = r.ReadBytes('E') }",
    scanner_default_with_buffer_grow => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"longline\")); sc.Buffer(make([]byte, 4), 128); for sc.Scan() { _ = sc.Text() } }",
    reader_read_string_then_peek => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"a,b\")); _, _ = r.ReadString(','); _, _ = r.Peek(1) }",
}

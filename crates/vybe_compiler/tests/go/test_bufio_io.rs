//! bufio and io packages: Scanner, Reader, Writer, ReadString, ReadBytes, ReadAll, Copy.

go_run_cases! {
    scanner_default_line_first_token => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"alpha\\nbeta\")); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["alpha"]
    ),
    scanner_default_line_second_token => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"alpha\\nbeta\")); sc.Scan(); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["beta"]
    ),
    scanner_scan_false_after_eof => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"x\")); sc.Scan(); fmt.Println(sc.Scan()) }",
        vec!["false"]
    ),
    scanner_bytes_matches_text => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"hi\")); sc.Scan(); fmt.Println(string(sc.Bytes()) == sc.Text()) }",
        vec!["true"]
    ),
    scanner_words_first_word => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"go lang\")); sc.Split(bufio.ScanWords); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["go"]
    ),
    scanner_words_second_word => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"go lang\")); sc.Split(bufio.ScanWords); sc.Scan(); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["lang"]
    ),
    scanner_bytes_one_byte => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"AB\")); sc.Split(bufio.ScanBytes); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["A"]
    ),
    scanner_runes_first_rune => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"日\")); sc.Split(bufio.ScanRunes); sc.Scan(); fmt.Println(sc.Text()) }",
        vec!["日"]
    ),
    scanner_empty_input_no_tokens => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"\")); fmt.Println(sc.Scan()) }",
        vec!["false"]
    ),
    scanner_count_three_lines => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"one\\ntwo\\nthree\\n\")); n := 0; for sc.Scan() { n++ }; fmt.Println(n) }",
        vec!["3"]
    ),

    reader_read_byte_first => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"xyz\")); b, _ := r.ReadByte(); fmt.Println(string(b)) }",
        vec!["x"]
    ),
    reader_unread_byte_rereads => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"go\")); b1, _ := r.ReadByte(); _ = b1; r.UnreadByte(); b2, _ := r.ReadByte(); fmt.Println(string(b2)) }",
        vec!["g"]
    ),
    reader_read_rune_unicode => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"日\")); ch, _, _ := r.ReadRune(); fmt.Println(string(ch)) }",
        vec!["日"]
    ),
    reader_peek_without_advancing => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"abc\")); peek, _ := r.Peek(2); fmt.Println(string(peek)); b, _ := r.ReadByte(); fmt.Println(string(b)) }",
        vec!["ab", "a"]
    ),
    reader_read_slice_until_comma => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"a,b\")); part, _ := r.ReadSlice(','); fmt.Println(string(part)) }",
        vec!["a,"]
    ),
    reader_read_string_until_newline => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"hi\\n\")); s, _ := r.ReadString('\\n'); fmt.Println(s) }",
        vec!["hi\n"]
    ),
    reader_read_bytes_until_semicolon => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"ok;rest\")); b, _ := r.ReadBytes(';'); fmt.Println(string(b)) }",
        vec!["ok;"]
    ),
    reader_readline_strips_newline => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"hello\\n\")); line, _, _ := r.ReadLine(); fmt.Println(string(line)) }",
        vec!["hello"]
    ),
    reader_buffered_after_prefetch => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReaderSize(strings.NewReader(\"abcd\"), 8); r.ReadByte(); fmt.Println(r.Buffered()) }",
        vec!["3"]
    ),
    reader_discard_skips_bytes => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"skipme\")); r.Discard(4); b, _ := r.ReadByte(); fmt.Println(string(b)) }",
        vec!["m"]
    ),
    reader_read_fills_buffer => (
        "package main; import \"fmt\"; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"go\")); buf := make([]byte, 10); n, _ := r.Read(buf); fmt.Println(n); fmt.Println(string(buf[:n])) }",
        vec!["2", "go"]
    ),

    writer_write_string_after_flush => (
        "package main; import \"fmt\"; import \"bufio\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := bufio.NewWriter(&buf); w.WriteString(\"vybe\"); w.Flush(); fmt.Println(buf.String()) }",
        vec!["vybe"]
    ),
    writer_write_byte => (
        "package main; import \"fmt\"; import \"bufio\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := bufio.NewWriter(&buf); w.WriteByte('Z'); w.Flush(); fmt.Println(buf.String()) }",
        vec!["Z"]
    ),
    writer_buffered_before_flush => (
        "package main; import \"fmt\"; import \"bufio\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := bufio.NewWriter(&buf); w.WriteString(\"ab\"); fmt.Println(w.Buffered()); w.Flush(); fmt.Println(buf.String()) }",
        vec!["2", "ab"]
    ),
    writer_reset_rebinds_output => (
        "package main; import \"fmt\"; import \"bufio\"; import \"bytes\"; func main() { var a bytes.Buffer; var b bytes.Buffer; w := bufio.NewWriter(&a); w.WriteString(\"old\"); w.Reset(&b); w.WriteString(\"new\"); w.Flush(); fmt.Println(a.String()); fmt.Println(b.String()) }",
        vec!["", "new"]
    ),

    io_readall_entire_content => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { data, _ := io.ReadAll(strings.NewReader(\"full\")); fmt.Println(string(data)) }",
        vec!["full"]
    ),
    io_readall_empty_reader => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { data, _ := io.ReadAll(strings.NewReader(\"\")); fmt.Println(len(data)) }",
        vec!["0"]
    ),
    io_copy_transfers_bytes => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; _, _ = io.Copy(&dst, strings.NewReader(\"copy\")); fmt.Println(dst.String()) }",
        vec!["copy"]
    ),
    io_copy_returns_byte_count => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; n, _ := io.Copy(&dst, strings.NewReader(\"four\")); fmt.Println(n) }",
        vec!["4"]
    ),
    io_readfull_exact_length => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { buf := make([]byte, 3); _, err := io.ReadFull(strings.NewReader(\"abc\"), buf); fmt.Println(string(buf)); fmt.Println(err == nil) }",
        vec!["abc", "true"]
    ),
    io_limitreader_truncates => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { lr := io.LimitReader(strings.NewReader(\"longtext\"), 4); data, _ := io.ReadAll(lr); fmt.Println(string(data)) }",
        vec!["long"]
    ),
    io_nopcloser_preserves_read => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { rc := io.NopCloser(strings.NewReader(\"wrap\")); data, _ := io.ReadAll(rc); fmt.Println(string(data)) }",
        vec!["wrap"]
    ),
}

go_compile_cases! {
    scanner_custom_split_on_pipe => "package main; import \"bufio\"; import \"strings\"; func main() { sc := bufio.NewScanner(strings.NewReader(\"a|b\")); sc.Split(func(data []byte, atEOF bool) (advance int, token []byte, err error) { for i, b := range data { if b == '|' { return i + 1, data[:i], nil } }; return 0, nil, nil }); _ = sc.Scan() }",
    reader_write_to_destination => "package main; import \"bufio\"; import \"bytes\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"go\")); var dst bytes.Buffer; _, _ = r.WriteTo(&dst) }",
    reader_reset_new_input => "package main; import \"bufio\"; import \"strings\"; func main() { r := bufio.NewReader(strings.NewReader(\"old\")); r.Reset(strings.NewReader(\"new\")) }",
    writer_write_rune => "package main; import \"bufio\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := bufio.NewWriter(&buf); _, _ = w.WriteRune('日'); w.Flush() }",
    writer_new_writer_size => "package main; import \"bufio\"; import \"bytes\"; func main() { var buf bytes.Buffer; w := bufio.NewWriterSize(&buf, 16); w.WriteString(\"size\") }",
    io_copy_buffer_custom => "package main; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; buf := make([]byte, 8); _, _ = io.CopyBuffer(&dst, strings.NewReader(\"buf\"), buf) }",
    io_multi_reader_sequential => "package main; import \"io\"; import \"strings\"; func main() { mr := io.MultiReader(strings.NewReader(\"ab\"), strings.NewReader(\"cd\")); _, _ = io.ReadAll(mr) }",
    io_multi_writer_duplicates => "package main; import \"io\"; import \"bytes\"; func main() { var a bytes.Buffer; var b bytes.Buffer; mw := io.MultiWriter(&a, &b); _, _ = mw.Write([]byte(\"x\")) }",
    io_pipe_roundtrip => "package main; import \"io\"; func main() { pr, pw := io.Pipe(); go func() { _, _ = pw.Write([]byte(\"p\")); pw.Close() }(); _, _ = io.ReadAll(pr) }",
}

//! io: Copy, CopyN, CopyBuffer, Pipe, TeeReader, ReadAll, ReadAtLeast,
//! WriteString via interface, LimitReader, MultiReader — extended runtime and
//! compile coverage distinct from `test_bufio_io.rs` and `test_io_fs_extended.rs`.

go_run_cases! {
    copy_empty_source_writes_nothing => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; n, err := io.Copy(&dst, strings.NewReader(\"\")); fmt.Println(n, err == nil, dst.Len()) }",
        vec!["0 true 0"]
    ),
    copy_single_byte_payload => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; n, _ := io.Copy(&dst, strings.NewReader(\"x\")); fmt.Println(n, dst.String()) }",
        vec!["1 x"]
    ),
    copy_unicode_rune_payload => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; _, _ = io.Copy(&dst, strings.NewReader(\"日\")); fmt.Println(dst.String()) }",
        vec!["日"]
    ),
    copy_to_discard_counts_bytes => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { n, _ := io.Copy(io.Discard, strings.NewReader(\"discard\")); fmt.Println(n) }",
        vec!["7"]
    ),
    copy_repeated_from_limited_reader => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; src := io.LimitReader(strings.NewReader(\"abcdef\"), 3); n, _ := io.Copy(&dst, src); fmt.Println(n, dst.String()) }",
        vec!["3 abc"]
    ),
    copy_chained_multi_reader_source => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; src := io.MultiReader(strings.NewReader(\"ab\"), strings.NewReader(\"cd\")); _, _ = io.Copy(&dst, src); fmt.Println(dst.String()) }",
        vec!["abcd"]
    ),
    copy_preserves_internal_newlines => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; _, _ = io.Copy(&dst, strings.NewReader(\"a\\nb\")); fmt.Println(dst.String()) }",
        vec!["a\nb"]
    ),
    copy_from_bytes_reader_slice => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; func main() { var dst bytes.Buffer; _, _ = io.Copy(&dst, bytes.NewReader([]byte(\"buf\"))); fmt.Println(dst.String()) }",
        vec!["buf"]
    ),
    copy_n_zero_limit_copies_nothing => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; n, _ := io.CopyN(&dst, strings.NewReader(\"hello\"), 0); fmt.Println(n, dst.String()) }",
        vec!["0 "]
    ),
    copy_n_one_byte_from_string => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; n, _ := io.CopyN(&dst, strings.NewReader(\"go\"), 1); fmt.Println(n, dst.String()) }",
        vec!["1 g"]
    ),
    copy_n_exact_source_length => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; n, _ := io.CopyN(&dst, strings.NewReader(\"vybe\"), 4); fmt.Println(n, dst.String()) }",
        vec!["4 vybe"]
    ),
    copy_n_partial_when_source_shorter => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; n, err := io.CopyN(&dst, strings.NewReader(\"ab\"), 5); fmt.Println(n, dst.String(), err != nil) }",
        vec!["2 ab true"]
    ),
    copy_n_two_step_drain => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; src := strings.NewReader(\"abcd\"); n1, _ := io.CopyN(&dst, src, 2); n2, _ := io.Copy(&dst, src); fmt.Println(n1, n2, dst.String()) }",
        vec!["2 2 abcd"]
    ),
    copy_buffer_small_buffer_size => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; buf := make([]byte, 2); n, _ := io.CopyBuffer(&dst, strings.NewReader(\"abcd\"), buf); fmt.Println(n, dst.String()) }",
        vec!["4 abcd"]
    ),
    copy_buffer_large_buffer_size => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; buf := make([]byte, 64); n, _ := io.CopyBuffer(&dst, strings.NewReader(\"tiny\"), buf); fmt.Println(n, dst.String()) }",
        vec!["4 tiny"]
    ),
    copy_buffer_reuses_same_slice => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; buf := make([]byte, 3); _, _ = io.CopyBuffer(&dst, strings.NewReader(\"xyz\"), buf); fmt.Println(len(buf), dst.String()) }",
        vec!["3 xyz"]
    ),
    read_all_whitespace_only => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { data, _ := io.ReadAll(strings.NewReader(\"   \")); fmt.Println(len(data), string(data)) }",
        vec!["3    "]
    ),
    read_all_binary_null_byte => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { data, _ := io.ReadAll(strings.NewReader(\"\\x00a\")); fmt.Println(len(data), int(data[0])) }",
        vec!["2 0"]
    ),
    read_all_from_limit_reader => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { data, _ := io.ReadAll(io.LimitReader(strings.NewReader(\"longer\"), 3)); fmt.Println(string(data)) }",
        vec!["lon"]
    ),
    read_all_from_multi_reader => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { mr := io.MultiReader(strings.NewReader(\"12\"), strings.NewReader(\"34\")); data, _ := io.ReadAll(mr); fmt.Println(string(data)) }",
        vec!["1234"]
    ),
    read_at_least_minimum_one_byte => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { buf := make([]byte, 4); n, err := io.ReadAtLeast(strings.NewReader(\"go\"), buf, 1); fmt.Println(n, string(buf[:n]), err == nil) }",
        vec!["2 go true"]
    ),
    read_at_least_exact_available => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { buf := make([]byte, 3); n, err := io.ReadAtLeast(strings.NewReader(\"abc\"), buf, 3); fmt.Println(n, string(buf), err == nil) }",
        vec!["3 abc true"]
    ),
    read_at_least_short_source_errors => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { buf := make([]byte, 5); n, err := io.ReadAtLeast(strings.NewReader(\"ab\"), buf, 4); fmt.Println(n, err != nil) }",
        vec!["2 true"]
    ),
    read_at_least_zero_minimum => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { buf := make([]byte, 2); n, err := io.ReadAtLeast(strings.NewReader(\"z\"), buf, 0); fmt.Println(n, string(buf[:n]), err == nil) }",
        vec!["1 z true"]
    ),
    write_string_via_writer_interface => (
        "package main; import \"fmt\"; import \"bytes\"; import \"io\"; func main() { var buf bytes.Buffer; var w io.Writer = &buf; n, _ := io.WriteString(w, \"iface\"); fmt.Println(n, buf.String()) }",
        vec!["5 iface"]
    ),
    write_string_empty_payload => (
        "package main; import \"fmt\"; import \"bytes\"; import \"io\"; func main() { var buf bytes.Buffer; n, _ := io.WriteString(&buf, \"\"); fmt.Println(n, buf.Len()) }",
        vec!["0 0"]
    ),
    write_string_unicode_via_interface => (
        "package main; import \"fmt\"; import \"bytes\"; import \"io\"; func main() { var buf bytes.Buffer; var w io.Writer = &buf; _, _ = io.WriteString(w, \"日\"); fmt.Println(buf.String()) }",
        vec!["日"]
    ),
    limit_reader_zero_returns_empty => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { data, _ := io.ReadAll(io.LimitReader(strings.NewReader(\"data\"), 0)); fmt.Println(len(data)) }",
        vec!["0"]
    ),
    limit_reader_one_byte_cap => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { data, _ := io.ReadAll(io.LimitReader(strings.NewReader(\"hello\"), 1)); fmt.Println(string(data)) }",
        vec!["h"]
    ),
    limit_reader_exact_boundary => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { lr := io.LimitReader(strings.NewReader(\"abcde\"), 5); data, _ := io.ReadAll(lr); fmt.Println(len(data), string(data)) }",
        vec!["5 abcde"]
    ),
    limit_reader_stops_before_extra_data => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { lr := io.LimitReader(strings.NewReader(\"abcdef\"), 4); data, _ := io.ReadAll(lr); fmt.Println(string(data)) }",
        vec!["abcd"]
    ),
    limit_reader_second_read_empty => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { lr := io.LimitReader(strings.NewReader(\"xy\"), 2); buf := make([]byte, 1); n1, _ := lr.Read(buf); n2, _ := lr.Read(buf); fmt.Println(n1, string(buf[:n1]), n2) }",
        vec!["1 y 1"]
    ),
    multi_reader_two_empty_leading => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { mr := io.MultiReader(strings.NewReader(\"\"), strings.NewReader(\"ok\")); data, _ := io.ReadAll(mr); fmt.Println(string(data)) }",
        vec!["ok"]
    ),
    multi_reader_three_segments => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { mr := io.MultiReader(strings.NewReader(\"a\"), strings.NewReader(\"b\"), strings.NewReader(\"c\")); data, _ := io.ReadAll(mr); fmt.Println(string(data)) }",
        vec!["abc"]
    ),
    multi_reader_single_byte_reads => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { mr := io.MultiReader(strings.NewReader(\"12\"), strings.NewReader(\"34\")); buf := make([]byte, 1); mr.Read(buf); fmt.Println(string(buf)); mr.Read(buf); fmt.Println(string(buf)) }",
        vec!["1", "2"]
    ),
    tee_reader_copies_to_buffer => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var side bytes.Buffer; tr := io.TeeReader(strings.NewReader(\"tee\"), &side); data, _ := io.ReadAll(tr); fmt.Println(string(data), side.String()) }",
        vec!["tee tee"]
    ),
    tee_reader_side_empty_until_read => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var side bytes.Buffer; tr := io.TeeReader(strings.NewReader(\"x\"), &side); fmt.Println(side.Len()); _, _ = io.ReadAll(tr); fmt.Println(side.String()) }",
        vec!["0", "x"]
    ),
    tee_reader_partial_read_duplicates_prefix => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var side bytes.Buffer; tr := io.TeeReader(strings.NewReader(\"abcd\"), &side); buf := make([]byte, 2); tr.Read(buf); fmt.Println(string(buf), side.String()) }",
        vec!["ab ab"]
    ),
    tee_reader_with_limit_reader => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var side bytes.Buffer; src := io.LimitReader(strings.NewReader(\"long\"), 2); tr := io.TeeReader(src, &side); data, _ := io.ReadAll(tr); fmt.Println(string(data), side.String()) }",
        vec!["lo lo"]
    ),
    copy_n_from_tee_reader => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var side bytes.Buffer; tr := io.TeeReader(strings.NewReader(\"go\"), &side); var dst bytes.Buffer; n, _ := io.CopyN(&dst, tr, 1); fmt.Println(n, dst.String(), side.String()) }",
        vec!["1 g g"]
    ),
    multi_reader_then_copy_n => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { mr := io.MultiReader(strings.NewReader(\"ab\"), strings.NewReader(\"cd\")); var dst bytes.Buffer; n, _ := io.CopyN(&dst, mr, 3); fmt.Println(n, dst.String()) }",
        vec!["3 abc"]
    ),
    read_full_vs_read_at_least_same_buffer => (
        "package main; import \"fmt\"; import \"io\"; import \"strings\"; func main() { buf1 := make([]byte, 2); buf2 := make([]byte, 2); _, e1 := io.ReadFull(strings.NewReader(\"xy\"), buf1); _, e2 := io.ReadAtLeast(strings.NewReader(\"xy\"), buf2, 2); fmt.Println(string(buf1), string(buf2), e1 == nil, e2 == nil) }",
        vec!["xy xy true true"]
    ),
    copy_buffer_from_multi_reader => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; mr := io.MultiReader(strings.NewReader(\"v\"), strings.NewReader(\"y\")); buf := make([]byte, 1); _, _ = io.CopyBuffer(&dst, mr, buf); fmt.Println(dst.String()) }",
        vec!["vy"]
    ),
    write_string_then_copy_back => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; func main() { var a bytes.Buffer; io.WriteString(&a, \"src\"); var b bytes.Buffer; _, _ = io.Copy(&b, &a); fmt.Println(b.String()) }",
        vec!["src"]
    ),
    limit_reader_with_copy_n => (
        "package main; import \"fmt\"; import \"io\"; import \"bytes\"; import \"strings\"; func main() { lr := io.LimitReader(strings.NewReader(\"abcdef\"), 3); var dst bytes.Buffer; n, _ := io.CopyN(&dst, lr, 2); fmt.Println(n, dst.String()) }",
        vec!["2 ab"]
    ),
}

go_compile_cases! {
    io_pipe_create_endpoints => "package main; import \"io\"; func main() { pr, pw := io.Pipe(); _ = pr; _ = pw }",
    io_pipe_write_read_goroutine => "package main; import \"io\"; func main() { pr, pw := io.Pipe(); go func() { _, _ = pw.Write([]byte(\"pipe\")); pw.Close() }(); _, _ = io.ReadAll(pr) }",
    io_pipe_close_writer => "package main; import \"io\"; func main() { pr, pw := io.Pipe(); _ = pw.Close(); _, _ = pr.Read(make([]byte, 1)) }",
    io_pipe_close_reader => "package main; import \"io\"; func main() { pr, pw := io.Pipe(); _ = pr.Close(); _, _ = pw.Write([]byte(\"x\")) }",
    io_copy_n_eof_error => "package main; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; _, _ = io.CopyN(&dst, strings.NewReader(\"a\"), 3) }",
    io_copy_buffer_nil_panics_compile => "package main; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var dst bytes.Buffer; _, _ = io.CopyBuffer(&dst, strings.NewReader(\"x\"), nil) }",
    io_multi_writer_three_sinks => "package main; import \"io\"; import \"bytes\"; func main() { var a bytes.Buffer; var b bytes.Buffer; var c bytes.Buffer; mw := io.MultiWriter(&a, &b, &c); _, _ = mw.Write([]byte(\"z\")) }",
    io_multi_reader_four_segments => "package main; import \"io\"; import \"strings\"; func main() { mr := io.MultiReader(strings.NewReader(\"1\"), strings.NewReader(\"2\"), strings.NewReader(\"3\"), strings.NewReader(\"4\")); _, _ = io.ReadAll(mr) }",
    io_tee_reader_to_discard => "package main; import \"io\"; import \"strings\"; func main() { tr := io.TeeReader(strings.NewReader(\"d\"), io.Discard); _, _ = io.ReadAll(tr) }",
    io_tee_reader_nested_copy => "package main; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var side bytes.Buffer; tr := io.TeeReader(strings.NewReader(\"n\"), &side); var dst bytes.Buffer; _, _ = io.Copy(&dst, tr) }",
    io_limit_reader_negative_compile => "package main; import \"io\"; import \"strings\"; func main() { _ = io.LimitReader(strings.NewReader(\"a\"), -1) }",
    io_read_at_least_into_subslice => "package main; import \"io\"; import \"strings\"; func main() { buf := make([]byte, 4); _, _ = io.ReadAtLeast(strings.NewReader(\"go\"), buf[:2], 2) }",
    io_write_string_to_multi_writer => "package main; import \"io\"; import \"bytes\"; func main() { var a bytes.Buffer; var b bytes.Buffer; mw := io.MultiWriter(&a, &b); _, _ = io.WriteString(mw, \"mw\") }",
    io_copy_from_section_reader => "package main; import \"io\"; import \"bytes\"; import \"strings\"; func main() { sr := strings.NewReader(\"section\"); var dst bytes.Buffer; _, _ = io.Copy(&dst, io.LimitReader(sr, 4)) }",
    io_pipe_copy_buffer_roundtrip => "package main; import \"io\"; import \"strings\"; func main() { pr, pw := io.Pipe(); go func() { buf := make([]byte, 8); _, _ = io.CopyBuffer(pw, strings.NewReader(\"buf\"), buf); pw.Close() }(); _, _ = io.ReadAll(pr) }",
    io_read_all_from_tee_reader => "package main; import \"io\"; import \"bytes\"; import \"strings\"; func main() { var side bytes.Buffer; tr := io.TeeReader(strings.NewReader(\"all\"), &side); _, _ = io.ReadAll(tr) }",
}

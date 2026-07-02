//! bytes.Buffer: Grow, ReadFrom, WriteTo, UnreadByte/Rune, Next, Truncate, Reset,
//! Read/Write variants, Equal, Bytes — distinct from `test_bytes_package.rs` (package-level
//! Compare/Contains/Index/Trim/Join) and incidental Buffer use in `test_bufio_io.rs` /
//! `test_io_fs_extended.rs` (io/bufio wrappers, not Buffer method semantics).


go_run_cases! {
    new_buffer_string_seed => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { b := bytes.NewBufferString(\"seed\"); fmt.Println(b.String()) }",
        vec!["seed"]
    ),
    new_buffer_slice_initial => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { b := bytes.NewBuffer([]byte(\"init\")); fmt.Println(b.String()) }",
        vec!["init"]
    ),
    new_buffer_nil_starts_empty => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { b := bytes.NewBuffer(nil); fmt.Println(b.Len()) }",
        vec!["0"]
    ),

    write_string_appends_text => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"go\"); fmt.Println(b.String()) }",
        vec!["go"]
    ),
    write_byte_single_ascii => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteByte('Z'); fmt.Println(b.String()) }",
        vec!["Z"]
    ),
    write_rune_multibyte_utf8 => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteRune('日'); fmt.Println(b.String()) }",
        vec!["日"]
    ),
    write_byte_slice_payload => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.Write([]byte(\"xy\")); fmt.Println(b.String()) }",
        vec!["xy"]
    ),
    write_sequential_concatenates => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"vy\"); b.WriteString(\"be\"); fmt.Println(b.String()) }",
        vec!["vybe"]
    ),

    len_tracks_unread_portion => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"abc\"); fmt.Println(b.Len()) }",
        vec!["3"]
    ),
    cap_after_grow_meets_minimum => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.Grow(32); fmt.Println(b.Cap() >= 32) }",
        vec!["true"]
    ),
    string_reflects_written_bytes => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"buf\"); fmt.Println(b.String()) }",
        vec!["buf"]
    ),
    bytes_returns_unread_after_partial_read => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"abc\"); b.ReadByte(); fmt.Println(string(b.Bytes())) }",
        vec!["bc"]
    ),

    reset_clears_then_rebuilds => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"old\"); b.Reset(); b.WriteString(\"new\"); fmt.Println(b.String()) }",
        vec!["new"]
    ),
    truncate_keeps_prefix_bytes => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"hello\"); b.Truncate(3); fmt.Println(b.String()) }",
        vec!["hel"]
    ),
    truncate_to_zero_empties_buffer => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"x\"); b.Truncate(0); fmt.Println(b.Len()) }",
        vec!["0"]
    ),

    read_byte_consumes_first => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"go\"); ch, _ := b.ReadByte(); fmt.Println(string(ch)) }",
        vec!["g"]
    ),
    read_rune_decodes_first_rune => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"日lang\"); r, _, _ := b.ReadRune(); fmt.Println(string(r)) }",
        vec!["日"]
    ),
    read_into_fixed_slice_count => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"abcd\"); buf := make([]byte, 2); n, _ := b.Read(buf); fmt.Println(n); fmt.Println(string(buf)) }",
        vec!["2", "ab"]
    ),
    read_reduces_unread_len => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"abcd\"); buf := make([]byte, 2); _, _ = b.Read(buf); fmt.Println(b.Len()) }",
        vec!["2"]
    ),

    next_peels_prefix_chunk => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"abcdef\"); chunk := b.Next(2); fmt.Println(string(chunk)); fmt.Println(b.String()) }",
        vec!["ab", "cdef"]
    ),
    next_beyond_len_takes_remainder => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"hi\"); chunk := b.Next(10); fmt.Println(string(chunk)); fmt.Println(b.Len()) }",
        vec!["hi", "0"]
    ),

    unread_byte_backtracks_one_byte => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"go\"); b.ReadByte(); b.UnreadByte(); ch, _ := b.ReadByte(); fmt.Println(string(ch)) }",
        vec!["g"]
    ),
    unread_rune_backtracks_rune => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"日x\"); b.ReadRune(); b.UnreadRune(); r, _, _ := b.ReadRune(); fmt.Println(string(r)) }",
        vec!["日"]
    ),

    grow_reserves_write_space => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var b bytes.Buffer; b.Grow(64); b.WriteString(\"x\"); fmt.Println(b.Len()) }",
        vec!["1"]
    ),
    readfrom_appends_reader_payload => (
        "package main; import \"fmt\"; import \"bytes\"; import \"strings\"; func main() { var b bytes.Buffer; _, _ = b.ReadFrom(strings.NewReader(\"vybe\")); fmt.Println(b.String()) }",
        vec!["vybe"]
    ),
    readfrom_reports_byte_count => (
        "package main; import \"fmt\"; import \"bytes\"; import \"strings\"; func main() { var b bytes.Buffer; n, _ := b.ReadFrom(strings.NewReader(\"four\")); fmt.Println(n) }",
        vec!["4"]
    ),
    writeto_copies_unread_to_destination => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var src bytes.Buffer; src.WriteString(\"copy\"); var dst bytes.Buffer; _, _ = src.WriteTo(&dst); fmt.Println(dst.String()) }",
        vec!["copy"]
    ),
    writeto_reports_bytes_written => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var src bytes.Buffer; src.WriteString(\"data\"); var dst bytes.Buffer; n, _ := src.WriteTo(&dst); fmt.Println(n) }",
        vec!["4"]
    ),
    writeto_drains_source_unread => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var src bytes.Buffer; src.WriteString(\"go\"); var dst bytes.Buffer; _, _ = src.WriteTo(&dst); fmt.Println(src.Len()) }",
        vec!["0"]
    ),

    equal_same_unread_content => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var a bytes.Buffer; var b bytes.Buffer; a.WriteString(\"go\"); b.WriteString(\"go\"); fmt.Println(a.Equal(b)) }",
        vec!["true"]
    ),
    equal_differs_after_partial_read => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var a bytes.Buffer; var b bytes.Buffer; a.WriteString(\"go\"); b.WriteString(\"go\"); a.ReadByte(); fmt.Println(a.Equal(b)) }",
        vec!["false"]
    ),
}

go_compile_cases! {
    buffer_available_after_grow => "package main; import \"bytes\"; func main() { var b bytes.Buffer; b.Grow(16); _ = b.Available() }",
    buffer_readfrom_empty_reader => "package main; import \"bytes\"; import \"strings\"; func main() { var b bytes.Buffer; _, _ = b.ReadFrom(strings.NewReader(\"\")) }",
    buffer_next_zero_returns_empty => "package main; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"x\"); _ = b.Next(0) }",
    buffer_write_after_read_requires_reset => "package main; import \"bytes\"; func main() { var b bytes.Buffer; b.WriteString(\"hi\"); b.ReadByte(); b.Reset(); _, _ = b.Write([]byte(\"ok\")) }",
    buffer_read_byte_empty_eof => "package main; import \"bytes\"; func main() { var b bytes.Buffer; _, _ = b.ReadByte() }",
}

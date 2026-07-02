//! fmt Errorf, Sscanf, Fscanf, and Fprint-to-buffer — distinct from
//! `test_fmt_sprintf_verbs.rs` (Sprintf format verbs) and `test_errors_package.rs`
//! (errors.Is / As / Unwrap chains).


go_run_cases! {
    // Errorf — formatted error strings (not error-chain semantics)
    errorf_static_message => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"disk full\"); fmt.Println(err.Error()) }",
        vec!["disk full"]
    ),
    errorf_decimal_placeholder => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"status %d\", 503); fmt.Println(err.Error()) }",
        vec!["status 503"]
    ),
    errorf_string_and_int => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"%s:%d\", \"timeout\", 30); fmt.Println(err.Error()) }",
        vec!["timeout:30"]
    ),
    errorf_float_precision => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"ratio %.2f\", 0.125); fmt.Println(err.Error()) }",
        vec!["ratio 0.13"]
    ),
    errorf_bool_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"enabled=%t\", false); fmt.Println(err.Error()) }",
        vec!["enabled=false"]
    ),
    errorf_quoted_string_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"bad token %q\", \"\\n\"); fmt.Println(err.Error()) }",
        vec!["bad token \"\\n\""]
    ),
    errorf_hex_uppercase => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"addr %X\", 255); fmt.Println(err.Error()) }",
        vec!["addr FF"]
    ),
    errorf_default_verb_struct => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"field %v\", struct{ N int }{N: 9}); fmt.Println(err.Error()) }",
        vec!["field {9}"]
    ),

    // Sscanf — parse formatted text from strings
    sscanf_single_decimal => (
        "package main; import \"fmt\"; func main() { var n int; c, _ := fmt.Sscanf(\"42\", \"%d\", &n); fmt.Println(c, n) }",
        vec!["1 42"]
    ),
    sscanf_single_string => (
        "package main; import \"fmt\"; func main() { var s string; c, _ := fmt.Sscanf(\"hello\", \"%s\", &s); fmt.Println(c, s) }",
        vec!["1 hello"]
    ),
    sscanf_int_then_string => (
        "package main; import \"fmt\"; func main() { var n int; var s string; c, _ := fmt.Sscanf(\"7 go\", \"%d %s\", &n, &s); fmt.Println(c, n, s) }",
        vec!["2 7 go"]
    ),
    sscanf_float_value => (
        "package main; import \"fmt\"; func main() { var f float64; c, _ := fmt.Sscanf(\"3.14\", \"%f\", &f); fmt.Println(c, f) }",
        vec!["1 3.14"]
    ),
    sscanf_hex_integer => (
        "package main; import \"fmt\"; func main() { var n int; c, _ := fmt.Sscanf(\"ff\", \"%x\", &n); fmt.Println(c, n) }",
        vec!["1 255"]
    ),
    sscanf_quoted_string => (
        "package main; import \"fmt\"; func main() { var s string; c, _ := fmt.Sscanf(\"\\\"go\\\"\", \"%q\", &s); fmt.Println(c, s) }",
        vec!["1 go"]
    ),
    sscanf_bool_true => (
        "package main; import \"fmt\"; func main() { var ok bool; c, _ := fmt.Sscanf(\"true\", \"%t\", &ok); fmt.Println(c, ok) }",
        vec!["1 true"]
    ),
    sscanf_leading_whitespace_skipped => (
        "package main; import \"fmt\"; func main() { var n int; c, _ := fmt.Sscanf(\"  99\", \"%d\", &n); fmt.Println(c, n) }",
        vec!["1 99"]
    ),
    sscanf_negative_integer => (
        "package main; import \"fmt\"; func main() { var n int; c, _ := fmt.Sscanf(\"-12\", \"%d\", &n); fmt.Println(c, n) }",
        vec!["1 -12"]
    ),

    // Fscanf — parse formatted text from io.Reader
    fscanf_strings_reader_int => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var n int; c, _ := fmt.Fscanf(strings.NewReader(\"55\"), \"%d\", &n); fmt.Println(c, n) }",
        vec!["1 55"]
    ),
    fscanf_strings_reader_word => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var s string; c, _ := fmt.Fscanf(strings.NewReader(\"vybe\"), \"%s\", &s); fmt.Println(c, s) }",
        vec!["1 vybe"]
    ),
    fscanf_int_and_string_pair => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var n int; var s string; c, _ := fmt.Fscanf(strings.NewReader(\"3 ok\"), \"%d %s\", &n, &s); fmt.Println(c, n, s) }",
        vec!["2 3 ok"]
    ),
    fscanf_bytes_reader_float => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var f float64; c, _ := fmt.Fscanf(bytes.NewReader([]byte(\"2.5\")), \"%f\", &f); fmt.Println(c, f) }",
        vec!["1 2.5"]
    ),

    // Fprint / Fprintf / Fprintln into bytes.Buffer
    fprint_int_to_buffer => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var buf bytes.Buffer; fmt.Fprint(&buf, 42); fmt.Println(buf.String()) }",
        vec!["42"]
    ),
    fprint_multiple_values_spaced => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var buf bytes.Buffer; fmt.Fprint(&buf, \"a\", 1); fmt.Println(buf.String()) }",
        vec!["a1"]
    ),
    fprintf_formatted_to_buffer => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var buf bytes.Buffer; fmt.Fprintf(&buf, \"id=%d\", 7); fmt.Println(buf.String()) }",
        vec!["id=7"]
    ),
    fprintln_adds_newline => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var buf bytes.Buffer; fmt.Fprintln(&buf, \"go\"); fmt.Println(buf.String()) }",
        vec!["go\n"]
    ),
    fprint_string_and_bool => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { var buf bytes.Buffer; fmt.Fprint(&buf, \"ok=\", true); fmt.Println(buf.String()) }",
        vec!["ok=true"]
    ),
}

go_compile_cases! {
    errorf_octal_and_width => "package main; import \"fmt\"; func main() { _ = fmt.Errorf(\"#%04o\", 7) }",
    sscanf_scanln_stops_at_newline => "package main; import \"fmt\"; func main() { var s string; _, _ = fmt.Sscanln(\"line\\nrest\", &s) }",
    fscanf_scanln_from_reader => "package main; import \"fmt\"; import \"strings\"; func main() { var s string; _, _ = fmt.Fscanln(strings.NewReader(\"one two\"), &s) }",
    fscanf_scan_three_fields => "package main; import \"fmt\"; import \"strings\"; func main() { var a, b, c int; _, _ = fmt.Fscan(strings.NewReader(\"1 2 3\"), &a, &b, &c) }",
    fprintf_float_verb_to_buffer => "package main; import \"fmt\"; import \"bytes\"; func main() { var buf bytes.Buffer; _, _ = fmt.Fprintf(&buf, \"%.1f\", 1.25) }",
}

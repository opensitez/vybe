//! strconv: distinct parsing/formatting APIs beyond Itoa/Atoi in test_strings_advanced.


go_run_cases! {
    strconv_parse_int_base10 => ("package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.ParseInt(\"42\", 10, 64); fmt.Println(n) }", vec!["42"]),
    strconv_parse_int_base16 => ("package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.ParseInt(\"ff\", 16, 64); fmt.Println(n) }", vec!["255"]),
    strconv_parse_uint_base10 => ("package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.ParseUint(\"99\", 10, 64); fmt.Println(n) }", vec!["99"]),
    strconv_format_int_binary => ("package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatInt(5, 2)) }", vec!["101"]),
    strconv_format_int_hex => ("package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatInt(255, 16)) }", vec!["ff"]),
    strconv_format_bool_true => ("package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatBool(true)) }", vec!["true"]),
    strconv_format_bool_false => ("package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatBool(false)) }", vec!["false"]),
    strconv_parse_bool_true => ("package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseBool(\"true\"); fmt.Println(v) }", vec!["true"]),
    strconv_parse_bool_false => ("package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseBool(\"false\"); fmt.Println(v) }", vec!["false"]),
    strconv_quote_ascii => ("package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.Quote(\"hi\")) }", vec!["\"hi\""]),
    strconv_unquote_quoted => ("package main; import \"fmt\"; import \"strconv\"; func main() { s, _ := strconv.Unquote(`\"go\"`); fmt.Println(s) }", vec!["go"]),
    strconv_format_float_fixed => ("package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatFloat(1.5, 'f', 1, 64)) }", vec!["1.5"]),
    strconv_parse_float_decimal => ("package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseFloat(\"3.14\", 64); fmt.Println(v) }", vec!["3.14"]),
    strconv_atoi_negative => ("package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.Atoi(\"-7\"); fmt.Println(n) }", vec!["-7"]),
    strconv_itoa_zero => ("package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.Itoa(0)) }", vec!["0"]),
}

go_compile_cases! {
    strconv_append_int_slice => "package main; import \"strconv\"; func main() { b := []byte{}; _, _ = strconv.AppendInt(b, 42, 10) }",
    strconv_append_bool_slice => "package main; import \"strconv\"; func main() { b := []byte{}; _, _ = strconv.AppendBool(b, true) }",
    strconv_append_quote_slice => "package main; import \"strconv\"; func main() { b := []byte{}; _, _ = strconv.AppendQuote(b, \"x\") }",
    strconv_can_backquote_simple => "package main; import \"strconv\"; func main() { _ = strconv.CanBackquote(\"abc\") }",
    strconv_format_uint_octal => "package main; import \"fmt\"; import \"strconv\"; func main() { _ = strconv.FormatUint(8, 8) }",
    strconv_parse_float_hex => "package main; import \"strconv\"; func main() { _, _ = strconv.ParseFloat(\"0x1.8p0\", 64) }",
    strconv_quoted_with_char => "package main; import \"strconv\"; func main() { _ = strconv.QuoteRune('A') }",
    strconv_append_quote_rune => "package main; import \"strconv\"; func main() { b := []byte{}; _, _ = strconv.AppendQuoteRune(b, 'Z') }",
    strconv_format_complex => "package main; import \"strconv\"; func main() { _ = strconv.FormatComplex(1+2i, 'f', 2, 64) }",
    strconv_parse_complex => "package main; import \"strconv\"; func main() { _, _ = strconv.ParseComplex(\"1+2i\", 64) }",
}

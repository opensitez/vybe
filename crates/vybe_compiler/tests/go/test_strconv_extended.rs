//! strconv: Atoi/Itoa, ParseBool, ParseFloat/FormatFloat variants, Quote/Unquote,
//! AppendInt/AppendFloat, ParseUint bit sizes, CanBackquote — extended coverage
//! distinct from `test_strconv_package.rs`.


go_run_cases! {
    atoi_positive_decimal => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, err := strconv.Atoi(\"12345\"); fmt.Println(n); fmt.Println(err == nil) }",
        vec!["12345", "true"]
    ),
    atoi_negative_decimal => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.Atoi(\"-99\"); fmt.Println(n) }",
        vec!["-99"]
    ),
    atoi_zero_string => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.Atoi(\"0\"); fmt.Println(n) }",
        vec!["0"]
    ),
    atoi_leading_zeros => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.Atoi(\"007\"); fmt.Println(n) }",
        vec!["7"]
    ),
    atoi_invalid_returns_error => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { _, err := strconv.Atoi(\"12x\"); fmt.Println(err != nil) }",
        vec!["true"]
    ),
    atoi_empty_string_error => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { _, err := strconv.Atoi(\"\"); fmt.Println(err != nil) }",
        vec!["true"]
    ),
    itoa_positive => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.Itoa(42)) }",
        vec!["42"]
    ),
    itoa_negative => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.Itoa(-15)) }",
        vec!["-15"]
    ),
    itoa_large_int => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.Itoa(1000000)) }",
        vec!["1000000"]
    ),

    parse_bool_one => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseBool(\"1\"); fmt.Println(v) }",
        vec!["true"]
    ),
    parse_bool_zero => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseBool(\"0\"); fmt.Println(v) }",
        vec!["false"]
    ),
    parse_bool_upper_t => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseBool(\"T\"); fmt.Println(v) }",
        vec!["true"]
    ),
    parse_bool_upper_f => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseBool(\"F\"); fmt.Println(v) }",
        vec!["false"]
    ),
    parse_bool_lower_t => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseBool(\"t\"); fmt.Println(v) }",
        vec!["true"]
    ),
    parse_bool_mixed_case_true => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseBool(\"True\"); fmt.Println(v) }",
        vec!["true"]
    ),
    parse_bool_invalid_error => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { _, err := strconv.ParseBool(\"maybe\"); fmt.Println(err != nil) }",
        vec!["true"]
    ),

    parse_float_negative_exponent => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseFloat(\"1.5e-2\", 64); fmt.Println(v) }",
        vec!["0.015"]
    ),
    parse_float_positive_exponent => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseFloat(\"2e3\", 64); fmt.Println(v) }",
        vec!["2000"]
    ),
    parse_float_leading_plus => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseFloat(\"+3.5\", 64); fmt.Println(v) }",
        vec!["3.5"]
    ),
    parse_float32_bit_size => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseFloat(\"1.25\", 32); fmt.Println(v) }",
        vec!["1.25"]
    ),
    parse_float_inf_positive => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseFloat(\"Inf\", 64); fmt.Println(v > 1e308) }",
        vec!["true"]
    ),
    parse_float_inf_negative => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseFloat(\"-Inf\", 64); fmt.Println(v < -1e308) }",
        vec!["true"]
    ),
    parse_float_nan => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { v, _ := strconv.ParseFloat(\"NaN\", 64); fmt.Println(v != v) }",
        vec!["true"]
    ),

    format_float_scientific_lower => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatFloat(1234.5, 'e', 2, 64)) }",
        vec!["1.23e+03"]
    ),
    format_float_scientific_upper => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatFloat(0.0012, 'E', 1, 64)) }",
        vec!["1.2E-03"]
    ),
    format_float_general_short => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatFloat(3.14, 'g', 3, 64)) }",
        vec!["3.14"]
    ),
    format_float_general_upper => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatFloat(1000.0, 'G', 4, 64)) }",
        vec!["1000"]
    ),
    format_float_hex_prefix => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { s := strconv.FormatFloat(10.0, 'x', 0, 64); fmt.Println(len(s) > 2) }",
        vec!["true"]
    ),
    format_float_precision_zero => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.FormatFloat(7.0, 'f', 0, 64)) }",
        vec!["7"]
    ),

    quote_escapes_newline => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.Quote(\"a\\nb\")) }",
        vec!["\"a\\nb\""]
    ),
    quote_escapes_tab => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.Quote(\"a\\tb\")) }",
        vec!["\"a\\tb\""]
    ),
    quote_escapes_backslash => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.Quote(\"\\\\\")) }",
        vec!["\"\\\\\""]
    ),
    quote_empty_string => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.Quote(\"\")) }",
        vec!["\"\"\""]
    ),
    unquote_hex_escape => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { s, _ := strconv.Unquote(`\"\\x41\"`); fmt.Println(s) }",
        vec!["A"]
    ),
    unquote_unicode_escape => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { s, _ := strconv.Unquote(`\"\\u03BB\"`); fmt.Println(int([]rune(s)[0])) }",
        vec!["955"]
    ),
    unquote_octal_escape => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { s, _ := strconv.Unquote(`\"\\101\"`); fmt.Println(s) }",
        vec!["A"]
    ),

    append_int_decimal_to_empty => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { b := strconv.AppendInt([]byte{}, 99, 10); fmt.Println(string(b)) }",
        vec!["99"]
    ),
    append_int_hex_to_buffer => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { b := strconv.AppendInt([]byte(\"0x\"), 255, 16); fmt.Println(string(b)) }",
        vec!["0xff"]
    ),
    append_int_negative => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { b := strconv.AppendInt([]byte{}, -8, 10); fmt.Println(string(b)) }",
        vec!["-8"]
    ),
    append_uint_octal => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { b := strconv.AppendUint([]byte{}, 8, 8); fmt.Println(string(b)) }",
        vec!["10"]
    ),
    append_float_fixed => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { b := strconv.AppendFloat([]byte{}, 2.5, 'f', 1, 64); fmt.Println(string(b)) }",
        vec!["2.5"]
    ),
    append_float_scientific => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { b := strconv.AppendFloat([]byte{}, 100.0, 'e', 0, 64); fmt.Println(len(string(b)) > 0) }",
        vec!["true"]
    ),

    parse_uint_bitsize_8_max => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.ParseUint(\"255\", 10, 8); fmt.Println(n) }",
        vec!["255"]
    ),
    parse_uint_bitsize_16 => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.ParseUint(\"65535\", 10, 16); fmt.Println(n) }",
        vec!["65535"]
    ),
    parse_uint_bitsize_32 => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.ParseUint(\"4294967295\", 10, 32); fmt.Println(n) }",
        vec!["4294967295"]
    ),
    parse_uint_overflow_bitsize_8 => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { _, err := strconv.ParseUint(\"256\", 10, 8); fmt.Println(err != nil) }",
        vec!["true"]
    ),
    parse_int_bitsize_8_negative => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.ParseInt(\"-128\", 10, 8); fmt.Println(n) }",
        vec!["-128"]
    ),
    parse_int_bitsize_16 => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.ParseInt(\"-32768\", 10, 16); fmt.Println(n) }",
        vec!["-32768"]
    ),
    parse_int_base2 => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { n, _ := strconv.ParseInt(\"1010\", 2, 64); fmt.Println(n) }",
        vec!["10"]
    ),

    can_backquote_plain_ascii => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.CanBackquote(\"hello\")) }",
        vec!["true"]
    ),
    can_backquote_rejects_newline => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.CanBackquote(\"a\\nb\")) }",
        vec!["false"]
    ),
    can_backquote_rejects_backslash => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.CanBackquote(\"a\\\\b\")) }",
        vec!["false"]
    ),
    can_backquote_empty => (
        "package main; import \"fmt\"; import \"strconv\"; func main() { fmt.Println(strconv.CanBackquote(\"\")) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    strconv_format_uint_decimal => "package main; import \"strconv\"; func main() { _ = strconv.FormatUint(100, 10) }",
    strconv_format_uint_hex => "package main; import \"strconv\"; func main() { _ = strconv.FormatUint(255, 16) }",
    strconv_append_bool_true => "package main; import \"strconv\"; func main() { _, _ = strconv.AppendBool([]byte{}, true) }",
    strconv_append_bool_false => "package main; import \"strconv\"; func main() { _, _ = strconv.AppendBool([]byte{}, false) }",
    strconv_append_quote_to_buffer => "package main; import \"strconv\"; func main() { _, _ = strconv.AppendQuote([]byte(\"pre\"), \"post\") }",
    strconv_append_quote_rune => "package main; import \"strconv\"; func main() { _, _ = strconv.AppendQuoteRune([]byte{}, '世') }",
    strconv_append_quote_rune_to_ascii => "package main; import \"strconv\"; func main() { _, _ = strconv.AppendQuoteRuneToASCII([]byte{}, 'A') }",
    strconv_append_quote_to_ascii => "package main; import \"strconv\"; func main() { _, _ = strconv.AppendQuoteToASCII([]byte{}, \"hi\") }",
    strconv_quote_rune_greek => "package main; import \"strconv\"; func main() { _ = strconv.QuoteRune('λ') }",
    strconv_quote_rune_to_ascii => "package main; import \"strconv\"; func main() { _ = strconv.QuoteRuneToASCII('日') }",
    strconv_quote_to_ascii => "package main; import \"strconv\"; func main() { _ = strconv.QuoteToASCII(\"go\") }",
    strconv_unquote_char_hex => "package main; import \"strconv\"; func main() { _, _, _ = strconv.UnquoteChar(`\\x41`, `\\`) }",
    strconv_unquote_char_unicode => "package main; import \"strconv\"; func main() { _, _, _ = strconv.UnquoteChar(`\\u03BB`, `\\`) }",
    strconv_is_print_ascii => "package main; import \"strconv\"; func main() { _ = strconv.IsPrint('A') }",
    strconv_is_print_rejects_control => "package main; import \"strconv\"; func main() { _ = strconv.IsPrint('\\x01') }",
    strconv_is_graph_punct => "package main; import \"strconv\"; func main() { _ = strconv.IsGraphic('!') }",
    strconv_parse_float_hex_p754 => "package main; import \"strconv\"; func main() { _, _ = strconv.ParseFloat(\"0x1.fp+2\", 64) }",
    strconv_parse_float_underscores => "package main; import \"strconv\"; func main() { _, _ = strconv.ParseFloat(\"1_000.5\", 64) }",
    strconv_parse_int_underscores => "package main; import \"strconv\"; func main() { _, _ = strconv.ParseInt(\"1_000\", 10, 64) }",
    strconv_parse_uint_base16 => "package main; import \"strconv\"; func main() { _, _ = strconv.ParseUint(\"deadbeef\", 16, 64) }",
    strconv_parse_uint_base2 => "package main; import \"strconv\"; func main() { _, _ = strconv.ParseUint(\"1111\", 2, 64) }",
    strconv_parse_int_base36 => "package main; import \"strconv\"; func main() { _, _ = strconv.ParseInt(\"z\", 36, 64) }",
    strconv_format_int_base36 => "package main; import \"strconv\"; func main() { _ = strconv.FormatInt(35, 36) }",
    strconv_format_float_bits_32 => "package main; import \"strconv\"; func main() { _ = strconv.FormatFloat(1.0, 'f', 2, 32) }",
    strconv_format_float_inf => "package main; import \"strconv\"; func main() { _ = strconv.FormatFloat(1.0/0.0, 'f', 0, 64) }",
    strconv_format_float_nan => "package main; import \"strconv\"; func main() { _ = strconv.FormatFloat(0.0/0.0, 'g', 0, 64) }",
    strconv_atoi_whitespace_trim => "package main; import \"strconv\"; func main() { _, _ = strconv.Atoi(\"  42  \") }",
    strconv_parse_bool_whitespace => "package main; import \"strconv\"; func main() { _, _ = strconv.ParseBool(\"  true  \") }",
    strconv_unquote_backtick_string => "package main; import \"strconv\"; func main() { _, _ = strconv.Unquote(`hello`) }",
    strconv_can_backquote_unicode => "package main; import \"strconv\"; func main() { _ = strconv.CanBackquote(\"日本語\") }",
}

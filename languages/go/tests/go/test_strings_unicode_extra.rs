use crate::helpers::*;

macro_rules! go_compile_test {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

macro_rules! go_run_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = run_prints($src);
            assert_eq!(out, $expected);
        }
    };
}

macro_rules! run_cases {
    ($( $name:ident => ($src:expr, $expected:expr), )*) => {
        $( go_run_test!($name, $src, $expected); )*
    };
}

macro_rules! compile_cases {
    ($( $name:ident => $src:expr, )*) => {
        $( go_compile_test!($name, $src); )*
    };
}

run_cases! {
    string_concatenation_chain_runtime => ("package main; import \"fmt\"; func main() { text := \"vy\" + \"be\" + \"go\"; fmt.Println(text); }", vec!["vybego"]),
    raw_string_literal_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(`go\\nraw`); }", vec!["go\\nraw"]),
    interpreted_string_escape_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(\"a\\tb\"); }", vec!["a\tb"]),
    string_len_ascii_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(len(\"gopher\")); }", vec!["6"]),
    string_index_byte_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(\"vybe\"[1]); }", vec!["121"]),
    substring_prefix_runtime => ("package main; import \"fmt\"; func main() { text := \"gopher\"; fmt.Println(text[:3]); }", vec!["gop"]),
    substring_suffix_runtime => ("package main; import \"fmt\"; func main() { text := \"gopher\"; fmt.Println(text[3:]); }", vec!["her"]),
    substring_middle_runtime => ("package main; import \"fmt\"; func main() { text := \"gopher\"; fmt.Println(text[1:4]); }", vec!["oph"]),
    string_compare_equal_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(\"go\" == \"go\"); }", vec!["true"]),
    string_compare_order_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(\"ant\" < \"bee\"); }", vec!["true"]),
    string_append_via_plus_equal_runtime => ("package main; import \"fmt\"; func main() { text := \"go\"; text += \"lang\"; fmt.Println(text); }", vec!["golang"]),
    string_in_slice_runtime => ("package main; import \"fmt\"; func main() { values := []string{\"a\", \"b\", \"c\"}; fmt.Println(values[2]); }", vec!["c"]),
    string_in_map_runtime => ("package main; import \"fmt\"; func main() { values := map[string]string{\"lang\": \"go\"}; fmt.Println(values[\"lang\"]); }", vec!["go"]),
    empty_string_default_runtime => ("package main; import \"fmt\"; func main() { var text string; fmt.Println(text == \"\"); }", vec!["true"]),
    string_switch_runtime => ("package main; import \"fmt\"; func main() { text := \"go\"; switch text { case \"go\": fmt.Println(1); default: fmt.Println(0) } }", vec!["1"]),
    string_join_like_manual_runtime => ("package main; import \"fmt\"; func main() { left, right := \"vy\", \"be\"; fmt.Println(left + \"-\" + right); }", vec!["vy-be"]),
    rune_literal_numeric_runtime => ("package main; import \"fmt\"; func main() { fmt.Println('A'); }", vec!["65"]),
    string_shadowing_runtime => ("package main; import \"fmt\"; func main() { text := \"outer\"; { text := \"inner\"; fmt.Println(text) }; fmt.Println(text); }", vec!["inner", "outer"]),
    string_function_result_runtime => ("package main; import \"fmt\"; func label(v int) string { if v > 0 { return \"pos\" }; return \"zero\" }; func main() { fmt.Println(label(1)); }", vec!["pos"]),
    string_builder_with_if_runtime => ("package main; import \"fmt\"; func main() { text := \"go\"; if len(text) == 2 { text = text + \"pher\" }; fmt.Println(text); }", vec!["gopher"]),
    string_trim_like_slice_runtime => ("package main; import \"fmt\"; func main() { text := \"[go]\"; fmt.Println(text[1:3]); }", vec!["go"]),
    string_index_on_literal_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(\"abc\"[2]); }", vec!["99"]),
    string_lexicographic_compare_runtime => ("package main; import \"fmt\"; func main() { fmt.Println(\"cat\" > \"car\"); }", vec!["true"]),
    string_concat_with_number_string_runtime => ("package main; import \"fmt\"; func main() { count := \"3\"; fmt.Println(\"items:\" + count); }", vec!["items:3"]),
    string_zero_length_slice_runtime => ("package main; import \"fmt\"; func main() { text := \"go\"; fmt.Println(len(text[:0])); }", vec!["0"]),
}

compile_cases! {
    unicode_string_literal_compile => "package main; const greeting = \"héllo\"; func main() { _ = greeting }",
    rune_literal_compile => "package main; func main() { _ = '界' }",
    string_range_compile => "package main; func main() { for i, r := range \"go\" { _, _ = i, r } }",
    string_to_byte_slice_compile => "package main; func main() { values := []byte(\"go\"); _ = values }",
    byte_slice_to_string_compile => "package main; func main() { values := []byte{103, 111}; _ = string(values) }",
    rune_slice_from_string_compile => "package main; func main() { values := []rune(\"go\"); _ = values }",
    string_map_lookup_compile => "package main; func main() { values := map[string]string{\"a\": \"b\"}; _, _ = values[\"a\"] }",
    raw_string_with_backticks_compile => "package main; const sample = `a\\nb`; func main() { _ = sample }",
    interpreted_string_escape_compile => "package main; const sample = \"a\\tb\"; func main() { _ = sample }",
    string_concat_with_named_type_compile => "package main; type label string; func main() { var left label = \"go\"; _ = string(left) + \"lang\" }",
    rune_alias_compile => "package main; type myRune rune; func main() { var r myRune = 'a'; _ = r }",
    string_array_compile => "package main; func main() { values := [2]string{\"a\", \"b\"}; _ = values }",
    string_struct_field_compile => "package main; type holder struct { text string }; func main() { _ = holder{text: \"go\"} }",
    string_type_alias_map_compile => "package main; type label string; func main() { values := map[label]int{\"go\": 1}; _ = values }",
    string_switch_init_compile => "package main; func main() { switch text := \"go\"; text { case \"go\": } }",
    string_in_interface_compile => "package main; func main() { var value interface{} = \"go\"; _ = value }",
    string_compare_not_equal_compile => "package main; func main() { _ = (\"go\" != \"rust\") }",
    string_slice_empty_compile => "package main; func main() { text := \"go\"; _ = text[0:0] }",
    string_literal_in_const_compile => "package main; const a = \"go\"; const b = a + \"lang\"; func main() { _ = b }",
    string_default_case_compile => "package main; func main() { switch \"x\" { default: _ = 1 } }",
    unicode_rune_in_switch_compile => "package main; func main() { switch 'λ' { case 'λ': _ = 1 } }",
    string_len_on_raw_literal_compile => "package main; func main() { _ = len(`go`) }",
    string_range_blank_identifier_compile => "package main; func main() { for _, r := range \"go\" { _ = r } }",
    string_concat_in_return_compile => "package main; func label() string { return \"go\" + \"lang\" }; func main() { _ = label() }",
    string_as_map_key_compile => "package main; func main() { values := map[string]int{\"go\": 1}; _ = values }",
}

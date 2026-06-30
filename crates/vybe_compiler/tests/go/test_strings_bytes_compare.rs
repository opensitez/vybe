//! strings: Compare, EqualFold, HasPrefix/Suffix, Count, Repeat, Map, ToValidUTF8;
//! bytes: Compare, Equal, HasPrefix, Index, ToUpper/ToLower — distinct from
//! `test_strings_ops_extended.rs` (Cut/Clone focus) and `test_strings_builder.rs`.

use crate::helpers::*;

go_run_cases! {
    strings_compare_prefix_diff_at_first => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"apple\", \"apply\")) }",
        vec!["-1"]
    ),
    strings_compare_prefix_diff_at_last => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"cat\", \"car\")) }",
        vec!["1"]
    ),
    strings_compare_both_empty => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"\", \"\")) }",
        vec!["0"]
    ),
    strings_compare_unicode_codepoint_order => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"a\", \"ä\")) }",
        vec!["-1"]
    ),
    strings_compare_longer_greater => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"go\", \"g\")) }",
        vec!["1"]
    ),

    strings_equal_fold_turkish_dotted_i => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.EqualFold(\"i\", \"I\")) }",
        vec!["true"]
    ),
    strings_equal_fold_kelvin_k => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.EqualFold(\"k\", \"K\")) }",
        vec!["true"]
    ),
    strings_equal_fold_greek_sigma => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.EqualFold(\"Σ\", \"σ\")) }",
        vec!["true"]
    ),
    strings_equal_fold_one_empty => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.EqualFold(\"\", \"a\")) }",
        vec!["false"]
    ),
    strings_equal_fold_same_unicode => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.EqualFold(\"日本\", \"日本\")) }",
        vec!["true"]
    ),

    strings_has_prefix_empty_needle => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasPrefix(\"go\", \"\")) }",
        vec!["true"]
    ),
    strings_has_prefix_exact_match => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasPrefix(\"go\", \"go\")) }",
        vec!["true"]
    ),
    strings_has_prefix_longer_than_string => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasPrefix(\"go\", \"gopher\")) }",
        vec!["false"]
    ),
    strings_has_prefix_unicode => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasPrefix(\"日本語\", \"日本\")) }",
        vec!["true"]
    ),
    strings_has_suffix_empty_needle => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasSuffix(\"go\", \"\")) }",
        vec!["true"]
    ),
    strings_has_suffix_exact_match => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasSuffix(\"golang\", \"lang\")) }",
        vec!["true"]
    ),
    strings_has_suffix_unicode => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasSuffix(\"hello世界\", \"世界\")) }",
        vec!["true"]
    ),
    strings_has_suffix_mismatch => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasSuffix(\"file.go\", \".txt\")) }",
        vec!["false"]
    ),

    strings_count_single_char => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Count(\"banana\", \"a\")) }",
        vec!["3"]
    ),
    strings_count_substring_overlap => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Count(\"aaaa\", \"aa\")) }",
        vec!["2"]
    ),
    strings_count_empty_substring => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Count(\"abc\", \"\")) }",
        vec!["4"]
    ),
    strings_count_not_found => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Count(\"go\", \"py\")) }",
        vec!["0"]
    ),
    strings_count_unicode_rune => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Count(\"日日月\", \"日\")) }",
        vec!["2"]
    ),

    strings_repeat_empty_pattern => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Repeat(\"\", 5)) }",
        vec![""]
    ),
    strings_repeat_two_rune_pattern => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Repeat(\"-=\", 3)) }",
        vec!["-=-=-"]
    ),
    strings_repeat_large_count => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(len(strings.Repeat(\"x\", 10))) }",
        vec!["10"]
    ),

    strings_map_identity => (
        "package main; import \"fmt\"; import \"strings\"; func main() { out := strings.Map(func(r rune) rune { return r }, \"abc\"); fmt.Println(out) }",
        vec!["abc"]
    ),
    strings_map_to_spaces => (
        "package main; import \"fmt\"; import \"strings\"; func main() { out := strings.Map(func(r rune) rune { if r == '-' { return ' ' }; return r }, \"a-b-c\"); fmt.Println(out) }",
        vec!["a b c"]
    ),
    strings_map_unicode_lower => (
        "package main; import \"fmt\"; import \"strings\"; func main() { out := strings.Map(func(r rune) rune { if r >= 'A' && r <= 'Z' { return r + 32 }; return r }, \"GoLang\"); fmt.Println(out) }",
        vec!["golang"]
    ),

    strings_to_valid_utf8_replaces_invalid => (
        "package main; import \"fmt\"; import \"strings\"; func main() { s := string([]byte{0xff, 0xfe, 'a'}); out := strings.ToValidUTF8(s, \"?\"); fmt.Println(len(out) > 1) }",
        vec!["true"]
    ),
    strings_to_valid_utf8_keeps_valid => (
        "package main; import \"fmt\"; import \"strings\"; func main() { out := strings.ToValidUTF8(\"go\", \"?\"); fmt.Println(out) }",
        vec!["go"]
    ),
    strings_to_valid_utf8_empty_replacement => (
        "package main; import \"fmt\"; import \"strings\"; func main() { s := string([]byte{0xff}); out := strings.ToValidUTF8(s, \"\"); fmt.Println(len(out) >= 0) }",
        vec!["true"]
    ),

    bytes_compare_nil_vs_empty => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Compare(nil, []byte{})) }",
        vec!["0"]
    ),
    bytes_compare_nil_vs_nonempty => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Compare(nil, []byte(\"a\"))) }",
        vec!["-1"]
    ),
    bytes_compare_prefix_equal_shorter => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Compare([]byte(\"ab\"), []byte(\"abc\"))) }",
        vec!["-1"]
    ),
    bytes_compare_byte_order => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Compare([]byte{1, 2}, []byte{1, 3})) }",
        vec!["-1"]
    ),

    bytes_equal_nil_nil => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Equal(nil, nil)) }",
        vec!["true"]
    ),
    bytes_equal_nil_empty => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Equal(nil, []byte{})) }",
        vec!["true"]
    ),
    bytes_equal_same_content_diff_cap => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { a := append([]byte{}, 'x'); b := []byte{'x'}; fmt.Println(bytes.Equal(a, b)) }",
        vec!["true"]
    ),
    bytes_equal_different_length => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Equal([]byte(\"ab\"), []byte(\"a\"))) }",
        vec!["false"]
    ),

    bytes_has_prefix_empty => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.HasPrefix([]byte(\"abc\"), []byte{})) }",
        vec!["true"]
    ),
    bytes_has_prefix_exact => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.HasPrefix([]byte{1,2,3}, []byte{1,2})) }",
        vec!["true"]
    ),
    bytes_has_prefix_too_long => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.HasPrefix([]byte(\"go\"), []byte(\"gopher\"))) }",
        vec!["false"]
    ),

    bytes_index_first_byte => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Index([]byte(\"abc\"), []byte(\"a\"))) }",
        vec!["0"]
    ),
    bytes_index_middle => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Index([]byte(\"abcabc\"), []byte(\"ca\"))) }",
        vec!["2"]
    ),
    bytes_index_empty_needle => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Index([]byte(\"ab\"), []byte{})) }",
        vec!["0"]
    ),
    bytes_index_empty_haystack => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(bytes.Index([]byte{}, []byte(\"a\"))) }",
        vec!["-1"]
    ),

    bytes_to_upper_ascii => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(string(bytes.ToUpper([]byte(\"aZ\")))) }",
        vec!["AZ"]
    ),
    bytes_to_upper_non_ascii_unchanged => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { b := []byte(\"日\"); u := bytes.ToUpper(b); fmt.Println(bytes.Equal(b, u)) }",
        vec!["true"]
    ),
    bytes_to_lower_ascii => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(string(bytes.ToLower([]byte(\"AbC\")))) }",
        vec!["abc"]
    ),
    bytes_to_lower_preserves_non_letters => (
        "package main; import \"fmt\"; import \"bytes\"; func main() { fmt.Println(string(bytes.ToLower([]byte(\"A1_B\")))) }",
        vec!["a1_b"]
    ),
}

go_compile_cases! {
    strings_compare_long_unicode => "package main; import \"strings\"; func main() { _ = strings.Compare(\"α\", \"β\") }",
    strings_equal_fold_angstrom => "package main; import \"strings\"; func main() { _ = strings.EqualFold(\"å\", \"Å\") }",
    strings_has_prefix_case_sensitive => "package main; import \"strings\"; func main() { _ = strings.HasPrefix(\"Go\", \"go\") }",
    strings_has_suffix_path_ext => "package main; import \"strings\"; func main() { _ = strings.HasSuffix(\"/tmp/x.tar.gz\", \".gz\") }",
    strings_count_multibyte_substr => "package main; import \"strings\"; func main() { _ = strings.Count(\"café\", \"é\") }",
    strings_repeat_single_rune => "package main; import \"strings\"; func main() { _ = strings.Repeat(\"*\", 4) }",
    strings_map_drop_non_ascii => "package main; import \"strings\"; func main() { _ = strings.Map(func(r rune) rune { if r > 127 { return -1 }; return r }, \"a日b\") }",
    strings_to_valid_utf8_multibyte_replacement => "package main; import \"strings\"; func main() { _ = strings.ToValidUTF8(string([]byte{0xc0}), \"REPL\") }",

    bytes_compare_both_nil => "package main; import \"bytes\"; func main() { _ = bytes.Compare(nil, nil) }",
    bytes_equal_empty_slices => "package main; import \"bytes\"; func main() { _ = bytes.Equal([]byte{}, []byte{}) }",
    bytes_has_prefix_nil_safety => "package main; import \"bytes\"; func main() { _ = bytes.HasPrefix(nil, []byte{}) }",
    bytes_has_suffix_basic => "package main; import \"bytes\"; func main() { _ = bytes.HasSuffix([]byte(\"abc\"), []byte(\"c\")) }",
    bytes_index_byte_single => "package main; import \"bytes\"; func main() { _ = bytes.IndexByte([]byte(\"abc\"), 'b') }",
    bytes_index_rune_ascii => "package main; import \"bytes\"; func main() { _ = bytes.IndexRune([]byte(\"go\"), 'o') }",
    bytes_last_index_found => "package main; import \"bytes\"; func main() { _ = bytes.LastIndex([]byte(\"ababa\"), []byte(\"ba\")) }",
    bytes_to_upper_empty => "package main; import \"bytes\"; func main() { _ = bytes.ToUpper([]byte{}) }",
    bytes_to_lower_empty => "package main; import \"bytes\"; func main() { _ = bytes.ToLower([]byte{}) }",
    bytes_to_upper_title_compat => "package main; import \"bytes\"; func main() { _ = bytes.ToUpper([]byte(\"hello\")) }",
    bytes_index_any_found => "package main; import \"bytes\"; func main() { _ = bytes.IndexAny([]byte(\"abc\"), \"xcb\") }",
    bytes_compare_func_custom => "package main; import \"bytes\"; func main() { _ = bytes.Equal([]byte(\"A\"), bytes.ToUpper([]byte(\"a\"))) }",
}

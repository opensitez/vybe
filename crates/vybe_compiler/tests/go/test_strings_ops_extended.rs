//! strings package: Cut, ContainsFunc/Any/Rune, FieldsFunc, SplitAfter/N,
//! Index/LastIndex variants, ReplaceAll, Trim* — distinct from `test_strings.rs`
//! (basic Contains/Index/Split/Join), `test_strings_advanced.rs` (Count/Fields/
//! Repeat/Compare), and `test_strings_builder.rs` (Builder/Reader/Replacer/
//! Compare/EqualFold/Map/Repeat).


go_run_cases! {
    cut_separator_found => (
        "package main; import \"fmt\"; import \"strings\"; func main() { before, after, found := strings.Cut(\"hello,world\", \",\"); fmt.Println(before); fmt.Println(after); fmt.Println(found) }",
        vec!["hello", "world", "true"]
    ),
    cut_separator_missing => (
        "package main; import \"fmt\"; import \"strings\"; func main() { before, after, found := strings.Cut(\"gopher\", \",\"); fmt.Println(before); fmt.Println(after); fmt.Println(found) }",
        vec!["gopher", "", "false"]
    ),
    cut_prefix_strips_known_prefix => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rest, found := strings.CutPrefix(\"https://host\", \"https://\"); fmt.Println(rest); fmt.Println(found) }",
        vec!["host", "true"]
    ),
    cut_prefix_no_match_returns_original => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rest, found := strings.CutPrefix(\"ftp://host\", \"https://\"); fmt.Println(rest); fmt.Println(found) }",
        vec!["ftp://host", "false"]
    ),
    cut_suffix_strips_known_suffix => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rest, found := strings.CutSuffix(\"file.txt\", \".txt\"); fmt.Println(rest); fmt.Println(found) }",
        vec!["file", "true"]
    ),
    cut_suffix_no_match_returns_original => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rest, found := strings.CutSuffix(\"file.go\", \".txt\"); fmt.Println(rest); fmt.Println(found) }",
        vec!["file.go", "false"]
    ),

    contains_any_first_rune_in_set => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ContainsAny(\"gopher\", \"xyzp\")) }",
        vec!["true"]
    ),
    contains_any_no_rune_in_set => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ContainsAny(\"go\", \"xyz\")) }",
        vec!["false"]
    ),
    contains_rune_present => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ContainsRune(\"vybe\", 'y')) }",
        vec!["true"]
    ),
    contains_rune_absent => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ContainsRune(\"go\", 'z')) }",
        vec!["false"]
    ),
    contains_func_matches_vowel => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ContainsFunc(\"rhythm\", func(r rune) bool { return r == 'a' || r == 'e' || r == 'i' || r == 'o' || r == 'u' })) }",
        vec!["false"]
    ),
    contains_func_matches_digit => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ContainsFunc(\"go2\", func(r rune) bool { return r >= '0' && r <= '9' })) }",
        vec!["true"]
    ),

    fields_func_splits_on_punctuation => (
        "package main; import \"fmt\"; import \"strings\"; func main() { f := strings.FieldsFunc(\"  a,b  c\", func(r rune) bool { return r == ' ' || r == ',' }); fmt.Println(len(f)); fmt.Println(f[0]); fmt.Println(f[2]) }",
        vec!["3", "a", "c"]
    ),
    fields_func_empty_string => (
        "package main; import \"fmt\"; import \"strings\"; func main() { f := strings.FieldsFunc(\"   \", func(r rune) bool { return r == ' ' }); fmt.Println(len(f)) }",
        vec!["0"]
    ),

    split_after_keeps_separator_on_left => (
        "package main; import \"fmt\"; import \"strings\"; func main() { parts := strings.SplitAfter(\"a,b,c\", \",\"); fmt.Println(len(parts)); fmt.Println(parts[0]); fmt.Println(parts[1]) }",
        vec!["3", "a,", "b,"]
    ),
    split_after_n_limits_segments => (
        "package main; import \"fmt\"; import \"strings\"; func main() { parts := strings.SplitAfterN(\"x-y-z\", \"-\", 2); fmt.Println(len(parts)); fmt.Println(parts[0]); fmt.Println(parts[1]) }",
        vec!["2", "x-", "y-z"]
    ),
    split_n_groups_remainder => (
        "package main; import \"fmt\"; import \"strings\"; func main() { parts := strings.SplitN(\"a,b,c,d\", \",\", 2); fmt.Println(len(parts)); fmt.Println(parts[0]); fmt.Println(parts[1]) }",
        vec!["2", "a", "b,c,d"]
    ),
    split_n_unlimited_negative_one => (
        "package main; import \"fmt\"; import \"strings\"; func main() { parts := strings.SplitN(\"one:two:three\", \":\", -1); fmt.Println(len(parts)) }",
        vec!["3"]
    ),

    index_byte_found => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.IndexByte(\"gopher\", 'p')) }",
        vec!["2"]
    ),
    index_byte_not_found => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.IndexByte(\"go\", 'z')) }",
        vec!["-1"]
    ),
    index_any_first_of_charset => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.IndexAny(\"abcdef\", \"cf\")) }",
        vec!["2"]
    ),
    index_rune_unicode => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.IndexRune(\"go日lang\", '日')) }",
        vec!["2"]
    ),
    index_func_first_uppercase => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.IndexFunc(\"vybeGo\", func(r rune) bool { return r >= 'A' && r <= 'Z' })) }",
        vec!["4"]
    ),

    last_index_substring => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.LastIndex(\"banana\", \"na\")) }",
        vec!["4"]
    ),
    last_index_byte => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.LastIndexByte(\"gopher\", 'e')) }",
        vec!["4"]
    ),
    last_index_any => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.LastIndexAny(\"abc123\", \"19\")) }",
        vec!["6"]
    ),
    last_index_func_trailing_digit => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.LastIndexFunc(\"ab12\", func(r rune) bool { return r >= '0' && r <= '9' })) }",
        vec!["3"]
    ),

    replace_all_every_occurrence => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ReplaceAll(\"a.a.a\", \".\", \"-\")) }",
        vec!["a-a-a"]
    ),
    replace_all_no_match_passthrough => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ReplaceAll(\"go\", \"rust\", \"zig\")) }",
        vec!["go"]
    ),
    replace_limited_count => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Replace(\"aabbcc\", \"b\", \"x\", 1)) }",
        vec!["aaxbcc"]
    ),

    trim_custom_cutset_both_ends => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Trim(\"!!hello!!\", \"!\")) }",
        vec!["hello"]
    ),
    trim_left_only => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.TrimLeft(\"xxxgo\", \"x\")) }",
        vec!["go"]
    ),
    trim_right_only => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.TrimRight(\"goxxx\", \"x\")) }",
        vec!["go"]
    ),
    trim_prefix_removed => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rest, ok := strings.TrimPrefix(\"prefix:value\", \"prefix:\"); fmt.Println(rest); fmt.Println(ok) }",
        vec!["value", "true"]
    ),
    trim_suffix_removed => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rest, ok := strings.TrimSuffix(\"name.go\", \".go\"); fmt.Println(rest); fmt.Println(ok) }",
        vec!["name", "true"]
    ),
}

go_compile_cases! {
    strings_clone_independent => "package main; import \"strings\"; func main() { s := \"go\"; c := strings.Clone(s); _ = c }",
    index_func_no_match => "package main; import \"strings\"; func main() { _ = strings.IndexFunc(\"abc\", func(r rune) bool { return r == 'z' }) }",
    last_index_func_no_match => "package main; import \"strings\"; func main() { _ = strings.LastIndexFunc(\"abc\", func(r rune) bool { return r == 'z' }) }",
    split_after_empty_separator => "package main; import \"strings\"; func main() { _ = strings.SplitAfter(\"ab\", \"\") }",
    trim_prefix_no_match => "package main; import \"strings\"; func main() { _, _ = strings.TrimPrefix(\"go\", \"rust\") }",
    trim_suffix_no_match => "package main; import \"strings\"; func main() { _, _ = strings.TrimSuffix(\"go\", \"lang\") }",
}

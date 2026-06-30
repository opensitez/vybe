//! regexp advanced runtime: FindAllStringSubmatch, SubexpNames, ReplaceAllString $refs,
//! LiteralPrefix, NumSubexp, QuoteMeta — distinct from `test_regexp_package.rs`.

use crate::helpers::*;

go_run_cases! {
    regexp_find_all_submatch_two_groups => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d+)-(\\d+)`); m := re.FindAllStringSubmatch(\"a1-2 b3-4\", -1); fmt.Println(len(m)); fmt.Println(m[0][1]); fmt.Println(m[1][2]) }",
        vec!["2", "1", "4"]
    ),
    regexp_find_all_submatch_limit_one => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`x(\\d)`); m := re.FindAllStringSubmatch(\"x1x2x3\", 1); fmt.Println(len(m)); fmt.Println(m[0][1]) }",
        vec!["1", "1"]
    ),
    regexp_find_all_submatch_no_match_empty => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`z(\\d)`); m := re.FindAllStringSubmatch(\"abc\", -1); fmt.Println(m == nil); fmt.Println(len(m)) }",
        vec!["true", "0"]
    ),
    regexp_find_all_submatch_word_triple => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\w)(\\w)(\\w)`); m := re.FindAllStringSubmatch(\"go!\", -1); fmt.Println(len(m[0])); fmt.Println(m[0][3]) }",
        vec!["4", "o"]
    ),
    regexp_find_all_submatch_email_parts => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`([\\w.]+)@([\\w.]+)`); m := re.FindAllStringSubmatch(\"a@b.com c@d.org\", -1); fmt.Println(m[0][1]); fmt.Println(m[1][2]) }",
        vec!["a", "org"]
    ),
    regexp_find_all_submatch_date_iso => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d{4})-(\\d{2})-(\\d{2})`); m := re.FindAllStringSubmatch(\"2024-06-30\", -1); fmt.Println(m[0][1]); fmt.Println(m[0][3]) }",
        vec!["2024", "30"]
    ),
    regexp_find_all_submatch_hex_color => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})`); m := re.FindAllStringSubmatch(\"#aabbcc\", -1); fmt.Println(m[0][1]); fmt.Println(m[0][4]) }",
        vec!["aa", "cc"]
    ),
    regexp_find_all_submatch_alternation => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(cat|dog)`); m := re.FindAllStringSubmatch(\"cat dog cat\", -1); fmt.Println(len(m)); fmt.Println(m[1][1]) }",
        vec!["3", "dog"]
    ),
    regexp_find_all_submatch_optional_group => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`colou?r`); m := re.FindAllStringSubmatch(\"color colour\", -1); fmt.Println(len(m)); fmt.Println(m[1][0]) }",
        vec!["2", "colour"]
    ),
    regexp_find_all_submatch_non_capturing => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(?:ab)+`); m := re.FindAllStringSubmatch(\"abab\", -1); fmt.Println(len(m[0])); fmt.Println(m[0][0]) }",
        vec!["1", "abab"]
    ),
    regexp_find_all_submatch_anchored_start => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`^(\\d+)`); m := re.FindAllStringSubmatch(\"42 rest\", -1); fmt.Println(m[0][1]); fmt.Println(len(m)) }",
        vec!["42", "1"]
    ),
    regexp_find_all_submatch_word_boundary => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`\\b(\\w{2})\\b`); m := re.FindAllStringSubmatch(\"go is ok\", -1); fmt.Println(len(m)); fmt.Println(m[0][1]) }",
        vec!["3", "go"]
    ),
    regexp_find_all_submatch_char_class => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`([aeiou])`); m := re.FindAllStringSubmatch(\"hello\", -1); fmt.Println(len(m)); fmt.Println(m[0][1]) }",
        vec!["2", "e"]
    ),
    regexp_find_all_submatch_greedy_plus => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(a+)`); m := re.FindAllStringSubmatch(\"aaab\", -1); fmt.Println(m[0][1]) }",
        vec!["aaa"]
    ),
    regexp_find_all_submatch_lazy_star => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(a*?)b`); m := re.FindAllStringSubmatch(\"aaab\", -1); fmt.Println(m[0][1]) }",
        vec!["aaa"]
    ),
    regexp_find_all_submatch_phone_digits => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`\\((\\d{3})\\)`); m := re.FindAllStringSubmatch(\"(555) call\", -1); fmt.Println(m[0][1]) }",
        vec!["555"]
    ),
    regexp_find_all_submatch_path_segments => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`/([\\w]+)`); m := re.FindAllStringSubmatch(\"/api/v1/users\", -1); fmt.Println(len(m)); fmt.Println(m[2][1]) }",
        vec!["3", "users"]
    ),
    regexp_find_all_submatch_whole_match_index_zero => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d+)`); m := re.FindAllStringSubmatch(\"n7\", -1); fmt.Println(m[0][0]) }",
        vec!["7"]
    ),
    regexp_subexp_names_named_group => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(?P<year>\\d{4})`); names := re.SubexpNames(); fmt.Println(names[1]); fmt.Println(len(names)) }",
        vec!["year", "2"]
    ),
    regexp_subexp_names_two_named => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(?P<a>\\w)(?P<b>\\w)`); names := re.SubexpNames(); fmt.Println(names[1]); fmt.Println(names[2]) }",
        vec!["a", "b"]
    ),
    regexp_subexp_names_unnamed_empty_first => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d+)`); names := re.SubexpNames(); fmt.Println(names[0] == \"\"); fmt.Println(names[1] == \"\") }",
        vec!["true", "true"]
    ),
    regexp_subexp_names_mixed_named_unnamed => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(?P<id>\\d+)-(\\w+)`); names := re.SubexpNames(); fmt.Println(names[1]); fmt.Println(names[2] == \"\") }",
        vec!["id", "true"]
    ),
    regexp_replace_all_dollar_one => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\w+)@(\\w+)`); fmt.Println(re.ReplaceAllString(\"a@b\", \"$1 at $2\")) }",
        vec!["a at b"]
    ),
    regexp_replace_all_dollar_two_first => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\w)(\\w)`); fmt.Println(re.ReplaceAllString(\"ab\", \"$2$1\")) }",
        vec!["ba"]
    ),
    regexp_replace_all_dollar_zero_whole => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d+)`); fmt.Println(re.ReplaceAllString(\"x1y2\", \"[$0]\")) }",
        vec!["x[1]y[2]"]
    ),
    regexp_replace_all_dollar_literal => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`a`); fmt.Println(re.ReplaceAllString(\"a\", \"$$\")) }",
        vec!["$"]
    ),
    regexp_replace_all_dollar_name_ref => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(?P<x>\\d+)`); fmt.Println(re.ReplaceAllString(\"9\", \"n=$x\")) }",
        vec!["n=9"]
    ),
    regexp_replace_all_multiple_groups => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d{2})-(\\d{2})-(\\d{4})`); fmt.Println(re.ReplaceAllString(\"06-30-2024\", \"$3/$1/$2\")) }",
        vec!["2024/06/30"]
    ),
    regexp_replace_all_global_substitution => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\w)`); fmt.Println(re.ReplaceAllString(\"abc\", \"[$1]\")) }",
        vec!["[a][b][c]"]
    ),
    regexp_literal_prefix_fixed_text => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`prefix(\\d+)`); p, lit := re.LiteralPrefix(); fmt.Println(p); fmt.Println(lit) }",
        vec!["prefix", "true"]
    ),
    regexp_literal_prefix_no_literal => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`^\\d+`); p, lit := re.LiteralPrefix(); fmt.Println(p); fmt.Println(lit) }",
        vec!["", "false"]
    ),
    regexp_literal_prefix_empty_pattern => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(``); p, lit := re.LiteralPrefix(); fmt.Println(p); fmt.Println(lit) }",
        vec!["", "true"]
    ),
    regexp_literal_prefix_metachar_not_literal => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`a.b`); _, lit := re.LiteralPrefix(); fmt.Println(lit) }",
        vec!["false"]
    ),
    regexp_num_subexp_one_group => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(a)`); fmt.Println(re.NumSubexp()) }",
        vec!["1"]
    ),
    regexp_num_subexp_two_groups => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`(a)(b)`); fmt.Println(re.NumSubexp()) }",
        vec!["2"]
    ),
    regexp_num_subexp_zero_plain => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`abc`); fmt.Println(re.NumSubexp()) }",
        vec!["0"]
    ),
    regexp_num_subexp_nested_parens => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { re := regexp.MustCompile(`((a)(b))`); fmt.Println(re.NumSubexp()) }",
        vec!["3"]
    ),
    regexp_quote_meta_dot_star => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { fmt.Println(regexp.QuoteMeta(\"a.b*\")) }",
        vec!["a\\.b\\*"]
    ),
    regexp_quote_meta_brackets => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { fmt.Println(regexp.QuoteMeta(\"[a-z]?\")) }",
        vec!["\\[a-z\\]\\?"]
    ),
    regexp_quote_meta_backslash => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { fmt.Println(regexp.QuoteMeta(`\\`)) }",
        vec!["\\\\"]
    ),
    regexp_quote_meta_plain_unchanged => (
        "package main; import \"fmt\"; import \"regexp\"; func main() { fmt.Println(regexp.QuoteMeta(\"hello\")) }",
        vec!["hello"]
    ),
}

go_compile_cases! {
    regexp_must_compile_valid => "package main; import \"regexp\"; func main() { _ = regexp.MustCompile(`\\d+`) }",
    regexp_must_compile_named => "package main; import \"regexp\"; func main() { _ = regexp.MustCompile(`(?P<n>\\w+)`) }",
    regexp_append_replacement_dollar_one => "package main; import \"regexp\"; import \"bytes\"; func main() { re := regexp.MustCompile(`(\\w)`); dst := []byte{}; src := []byte(\"a\"); _ = re.ReplaceAllFunc(src, func(m []byte) []byte { return re.Expand(dst[:0], []byte(\"[$1]\"), src, re.FindIndex(src)) }) }",
    regexp_append_replacement_named => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`(?P<v>\\d+)`); b := re.Expand(nil, []byte(\"v=$v\"), []byte(\"42\"), re.FindSubmatchIndex([]byte(\"42\"))); _ = b }",
    regexp_expand_dst_buffer => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d+)`); _ = re.Expand([]byte{}, []byte(\"n=$1\"), []byte(\"7\"), re.FindSubmatchIndex([]byte(\"7\"))) }",
    regexp_replace_all_func_callback => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`a`); _ = re.ReplaceAllFunc([]byte(\"aba\"), func(b []byte) []byte { return []byte(\"X\") }) }",
    regexp_find_all_submatch_index => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d)`); _ = re.FindAllSubmatchIndex([]byte(\"a1b2\"), -1) }",
    regexp_find_submatch_index => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\w+)`); _ = re.FindSubmatchIndex([]byte(\"go\")) }",
    regexp_longest_prefix => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`foo`); _ = re.Longest() }",
    regexp_subexp_index_by_name => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`(?P<id>\\d+)`); _ = re.SubexpIndex(\"id\") }",
    regexp_compile_invalid_pattern => "package main; import \"regexp\"; func main() { _, err := regexp.Compile(`(`); _ = err }",
    regexp_match_reader => "package main; import \"regexp\"; import \"strings\"; func main() { re := regexp.MustCompile(`go`); _, _ = re.MatchReader(strings.NewReader(\"gopher\")) }",
    regexp_find_reader => "package main; import \"regexp\"; import \"strings\"; func main() { re := regexp.MustCompile(`(\\d+)`); _, _ = re.FindReaderSubmatchIndex(strings.NewReader(\"n42\")) }",
    regexp_replace_all_literal => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`x`); _ = re.ReplaceAllLiteral([]byte(\"x\"), []byte(\"y\")) }",
    regexp_replace_all_literal_string => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`a+`); _ = re.ReplaceAllLiteralString(\"aaa\", \"b\") }",
    regexp_split_after => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`\\s`); _ = re.SplitAfter(\"a b c\", -1) }",
    regexp_split_after_n => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`,`); _ = re.SplitAfterN(\"a,b,c\", 2) }",
    regexp_find_string_index => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`go`); _ = re.FindStringIndex(\"gopher\") }",
    regexp_find_all_string => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`\\d+`); _ = re.FindAllString(\"a1b22\", -1) }",
    regexp_find_all_string_n => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`x`); _ = re.FindAllString(\"xxx\", 2) }",
    regexp_find_string_submatch_index => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\w)`); _ = re.FindStringSubmatchIndex(\"a\") }",
    regexp_replace_all_byte_slice => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`(\\d)`); _ = re.ReplaceAll([]byte(\"a1\"), []byte(\"$1\")) }",
    regexp_equal_pattern => "package main; import \"regexp\"; func main() { a := regexp.MustCompile(`a`); b := regexp.MustCompile(`a`); _ = a.String() == b.String() }",
    regexp_string_returns_pattern => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`abc`); _ = re.String() }",
    regexp_copy_on_write => "package main; import \"regexp\"; func main() { re := regexp.MustCompile(`test`); _ = re.Copy() }",
}

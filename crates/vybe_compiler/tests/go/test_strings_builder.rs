//! strings.Builder, Reader, Replacer, Compare, EqualFold, Map, Repeat.

use crate::helpers::*;

go_run_cases! {
    builder_write_string => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var b strings.Builder; b.WriteString(\"go\"); fmt.Println(b.String()) }",
        vec!["go"]
    ),
    builder_write_byte => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var b strings.Builder; b.WriteByte('A'); fmt.Println(b.String()) }",
        vec!["A"]
    ),
    builder_write_rune_unicode => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var b strings.Builder; b.WriteRune('日'); fmt.Println(b.String()) }",
        vec!["日"]
    ),
    builder_concat_writes => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var b strings.Builder; b.WriteString(\"vy\"); b.WriteString(\"be\"); fmt.Println(b.String()) }",
        vec!["vybe"]
    ),
    builder_len_after_writes => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var b strings.Builder; b.WriteString(\"abc\"); fmt.Println(b.Len()) }",
        vec!["3"]
    ),
    builder_reset_then_rebuild => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var b strings.Builder; b.WriteString(\"old\"); b.Reset(); b.WriteString(\"new\"); fmt.Println(b.String()) }",
        vec!["new"]
    ),
    builder_write_byte_slice => (
        "package main; import \"fmt\"; import \"strings\"; func main() { var b strings.Builder; b.Write([]byte(\"xy\")); fmt.Println(b.String()) }",
        vec!["xy"]
    ),

    reader_len_tracks_unread => (
        "package main; import \"fmt\"; import \"strings\"; func main() { r := strings.NewReader(\"abc\"); fmt.Println(r.Len()) }",
        vec!["3"]
    ),
    reader_size_is_original_length => (
        "package main; import \"fmt\"; import \"strings\"; func main() { r := strings.NewReader(\"go\"); fmt.Println(r.Size()) }",
        vec!["2"]
    ),
    reader_read_byte_first => (
        "package main; import \"fmt\"; import \"strings\"; func main() { r := strings.NewReader(\"go\"); b, _ := r.ReadByte(); fmt.Println(string(b)) }",
        vec!["g"]
    ),
    reader_read_rune_first => (
        "package main; import \"fmt\"; import \"strings\"; func main() { r := strings.NewReader(\"日\"); ch, _, _ := r.ReadRune(); fmt.Println(string(ch)) }",
        vec!["日"]
    ),
    reader_read_into_buffer => (
        "package main; import \"fmt\"; import \"strings\"; func main() { r := strings.NewReader(\"abc\"); buf := make([]byte, 2); n, _ := r.Read(buf); fmt.Println(n); fmt.Println(string(buf)) }",
        vec!["2", "ab"]
    ),
    reader_unread_byte_reread => (
        "package main; import \"fmt\"; import \"strings\"; func main() { r := strings.NewReader(\"go\"); b1, _ := r.ReadByte(); _ = b1; r.UnreadByte(); b2, _ := r.ReadByte(); fmt.Println(string(b2)) }",
        vec!["g"]
    ),
    reader_seek_start => (
        "package main; import \"fmt\"; import \"strings\"; func main() { r := strings.NewReader(\"go\"); _, _ = r.ReadByte(); pos, _ := r.Seek(0, 0); fmt.Println(pos) }",
        vec!["0"]
    ),

    replacer_single_pair => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rep := strings.NewReplacer(\"a\", \"x\"); fmt.Println(rep.Replace(\"aabb\")) }",
        vec!["xxbb"]
    ),
    replacer_ordered_pairs => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rep := strings.NewReplacer(\"a\", \"b\", \"b\", \"c\"); fmt.Println(rep.Replace(\"ab\")) }",
        vec!["bc"]
    ),
    replacer_no_match_passthrough => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rep := strings.NewReplacer(\"z\", \"q\"); fmt.Println(rep.Replace(\"go\")) }",
        vec!["go"]
    ),
    replacer_write_string_count => (
        "package main; import \"fmt\"; import \"strings\"; func main() { rep := strings.NewReplacer(\"x\", \"y\"); var b strings.Builder; n, _ := rep.WriteString(&b, \"x\"); fmt.Println(n); fmt.Println(b.String()) }",
        vec!["1", "y"]
    ),

    compare_equal_returns_zero => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"go\", \"go\")) }",
        vec!["0"]
    ),
    compare_greater_returns_one => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"z\", \"a\")) }",
        vec!["1"]
    ),
    compare_empty_less_than_nonempty => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"\", \"a\")) }",
        vec!["-1"]
    ),
    compare_shorter_prefix_less => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"go\", \"gopher\")) }",
        vec!["-1"]
    ),

    equal_fold_ascii_insensitive => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.EqualFold(\"Go\", \"go\")) }",
        vec!["true"]
    ),
    equal_fold_ascii_mismatch => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.EqualFold(\"cat\", \"dog\")) }",
        vec!["false"]
    ),
    equal_fold_both_empty => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.EqualFold(\"\", \"\")) }",
        vec!["true"]
    ),
    equal_fold_length_mismatch => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.EqualFold(\"go\", \"golang\")) }",
        vec!["false"]
    ),

    map_uppercase_runes => (
        "package main; import \"fmt\"; import \"strings\"; func main() { out := strings.Map(func(r rune) rune { if r >= 'a' && r <= 'z' { return r - 32 }; return r }, \"AbC\"); fmt.Println(out) }",
        vec!["ABC"]
    ),
    map_drops_vowels => (
        "package main; import \"fmt\"; import \"strings\"; func main() { out := strings.Map(func(r rune) rune { if r == 'a' || r == 'e' || r == 'i' || r == 'o' || r == 'u' { return -1 }; return r }, \"goaeiu\"); fmt.Println(out) }",
        vec!["g"]
    ),
    map_masks_digits => (
        "package main; import \"fmt\"; import \"strings\"; func main() { out := strings.Map(func(r rune) rune { if r >= '0' && r <= '9' { return '#' }; return r }, \"a1b2\"); fmt.Println(out) }",
        vec!["a#b#"]
    ),

    repeat_zero_count_empty => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Repeat(\"x\", 0)) }",
        vec![""]
    ),
    repeat_one_returns_original => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Repeat(\"go\", 1)) }",
        vec!["go"]
    ),
    repeat_multi_char_pattern => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Repeat(\"ab\", 3)) }",
        vec!["ababab"]
    ),
    repeat_unicode_rune => (
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Repeat(\"日\", 2)) }",
        vec!["日日"]
    ),
}

go_compile_cases! {
    builder_grow_before_write => "package main; import \"strings\"; func main() { var b strings.Builder; b.Grow(8); b.WriteString(\"grow\") }",
    reader_seek_current_relative => "package main; import \"strings\"; func main() { r := strings.NewReader(\"go\"); _, _ = r.ReadByte(); _, _ = r.Seek(1, 1) }",
    reader_seek_end => "package main; import \"strings\"; func main() { r := strings.NewReader(\"go\"); _, _ = r.Seek(0, 2) }",
    reader_read_at_offset => "package main; import \"strings\"; func main() { r := strings.NewReader(\"abc\"); buf := make([]byte, 1); _, _ = r.ReadAt(buf, 1) }",
    reader_unread_rune => "package main; import \"strings\"; func main() { r := strings.NewReader(\"日\"); _, _, _ = r.ReadRune(); _ = r.UnreadRune() }",
    reader_write_to => "package main; import \"strings\"; func main() { r := strings.NewReader(\"go\"); var b strings.Builder; _, _ = r.WriteTo(&b) }",
    equal_fold_unicode_sigma => "package main; import \"strings\"; func main() { _ = strings.EqualFold(\"σ\", \"Σ\") }",
    map_drop_all_runes => "package main; import \"strings\"; func main() { _ = strings.Map(func(r rune) rune { return -1 }, \"abc\") }",
}

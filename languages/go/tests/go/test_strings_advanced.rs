use crate::helpers::*;

#[test]
fn string_raw_literal() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := `line1\nline2`; fmt.Println(len(s)); }",
    );
    assert_eq!(out, vec!["11"]);
}
#[test]
fn string_escape_newline() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"line1\\nline2\"; fmt.Println(len(s)); }",
    );
    assert_eq!(out, vec!["11"]);
}
#[test]
fn string_escape_tab() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"a\\tb\"; fmt.Println(len(s)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn string_escape_quotes() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"\\\"hello\\\"\"; fmt.Println(s); }",
    );
    assert_eq!(out, vec!["\"hello\""]);
}
#[test]
fn string_index_byte() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"abc\"; fmt.Println(s[1]); }",
    );
    assert_eq!(out, vec!["98"]); // ascii for 'b'
}
#[test]
fn string_slice() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"hello world\"; fmt.Println(s[0:5]); }",
    );
    assert_eq!(out, vec!["hello"]);
}
#[test]
fn string_slice_start() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"hello world\"; fmt.Println(s[:5]); }",
    );
    assert_eq!(out, vec!["hello"]);
}
#[test]
fn string_slice_end() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"hello world\"; fmt.Println(s[6:]); }",
    );
    assert_eq!(out, vec!["world"]);
}
#[test]
fn string_iterate_bytes() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"abc\"; for i := 0; i < len(s); i++ { fmt.Println(s[i]) } }",
    );
    assert_eq!(out, vec!["97", "98", "99"]);
}
#[test]
fn string_iterate_range() {
    compile_ok(
        "package main; import \"fmt\"; func main() { s := \"abc\"; for i, r := range s { _ = i; _ = r; } }",
    );
}
#[test]
fn string_to_byte_slice() {
    compile_ok("package main; import \"fmt\"; func main() { s := \"abc\"; b := []byte(s); _ = b }");
}
#[test]
fn byte_slice_to_string() {
    compile_ok(
        "package main; import \"fmt\"; func main() { b := []byte{97, 98, 99}; s := string(b); _ = s }",
    );
}
#[test]
fn string_to_rune_slice() {
    compile_ok("package main; import \"fmt\"; func main() { s := \"abc\"; r := []rune(s); _ = r }");
}
#[test]
fn rune_slice_to_string() {
    compile_ok(
        "package main; import \"fmt\"; func main() { r := []rune{97, 98, 99}; s := string(r); _ = s }",
    );
}
#[test]
fn int_to_string() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strconv\"; func main() { s := strconv.Itoa(42); fmt.Println(s); }",
    );
    assert_eq!(out, vec!["42"]);
}
#[test]
fn string_to_int() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strconv\"; func main() { i, _ := strconv.Atoi(\"42\"); fmt.Println(i); }",
    );
    assert_eq!(out, vec!["42"]);
}
#[test]
fn strings_count() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Count(\"cheese\", \"e\")); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn strings_fields() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { f := strings.Fields(\" a b  c \"); fmt.Println(len(f)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn strings_repeat() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Repeat(\"a\", 5)); }",
    );
    assert_eq!(out, vec!["aaaaa"]);
}
#[test]
fn strings_compare() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Compare(\"a\", \"b\")); }",
    );
    assert_eq!(out, vec!["-1"]);
}

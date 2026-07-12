use crate::helpers::*;

#[test]
fn string_concatenation() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a := \"hello\"; b := \" world\"; fmt.Println(a + b); }",
    );
    assert_eq!(out, vec!["hello world"]);
}
#[test]
fn string_len() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"hello\"; fmt.Println(len(s)); }",
    );
    assert_eq!(out, vec!["5"]);
}
#[test]
fn string_empty_len() {
    let out =
        run_prints("package main; import \"fmt\"; func main() { s := \"\"; fmt.Println(len(s)); }");
    assert_eq!(out, vec!["0"]);
}
#[test]
fn string_contains_true() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Contains(\"hello world\", \"world\")); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn string_contains_false() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Contains(\"hello\", \"xyz\")); }",
    );
    assert_eq!(out, vec!["false"]);
}
#[test]
fn string_to_upper() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ToUpper(\"hello\")); }",
    );
    assert_eq!(out, vec!["HELLO"]);
}
#[test]
fn string_to_lower() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.ToLower(\"WORLD\")); }",
    );
    assert_eq!(out, vec!["world"]);
}
#[test]
fn string_trim_space() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.TrimSpace(\"  hi  \")); }",
    );
    assert_eq!(out, vec!["hi"]);
}
#[test]
fn string_has_prefix_true() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasPrefix(\"golang\", \"go\")); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn string_has_prefix_false() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasPrefix(\"golang\", \"lang\")); }",
    );
    assert_eq!(out, vec!["false"]);
}
#[test]
fn string_has_suffix_true() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasSuffix(\"golang\", \"lang\")); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn string_has_suffix_false() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.HasSuffix(\"golang\", \"go\")); }",
    );
    assert_eq!(out, vec!["false"]);
}
#[test]
fn string_index_found() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Index(\"hello\", \"ll\")); }",
    );
    assert_eq!(out, vec!["2"]);
}
#[test]
fn string_index_not_found() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Index(\"hello\", \"xyz\")); }",
    );
    assert_eq!(out, vec!["-1"]);
}
#[test]
fn string_replace_all() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { fmt.Println(strings.Replace(\"aabbcc\", \"b\", \"x\", -1)); }",
    );
    assert_eq!(out, vec!["aaxxcc"]);
}
#[test]
fn string_equality_true() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { a := \"go\"; b := \"go\"; fmt.Println(a == b); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn string_equality_false() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { fmt.Println(\"go\" == \"rust\"); }",
    );
    assert_eq!(out, vec!["false"]);
}
#[test]
fn string_inequality() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { fmt.Println(\"go\" != \"rust\"); }",
    );
    assert_eq!(out, vec!["true"]);
}
#[test]
fn string_returned_from_func() {
    let out = run_prints(
        "package main; import \"fmt\"; func greet(name string) string { return \"Hello \" + name } func main() { fmt.Println(greet(\"Go\")); }",
    );
    assert_eq!(out, vec!["Hello Go"]);
}
#[test]
fn string_passed_to_function() {
    let out = run_prints(
        "package main; import \"fmt\"; func shout(s string) { fmt.Println(s + \"!\") } func main() { shout(\"hello\"); }",
    );
    assert_eq!(out, vec!["hello!"]);
}
#[test]
fn string_concat_in_condition() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { prefix := \"he\"; s := \"hello\"; if prefix + \"llo\" == s { fmt.Println(\"match\"); } }",
    );
    assert_eq!(out, vec!["match"]);
}
#[test]
fn string_len_condition() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := \"abc\"; if len(s) > 2 { fmt.Println(\"long\"); } }",
    );
    assert_eq!(out, vec!["long"]);
}
#[test]
fn multiple_strings_println() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { fmt.Println(\"one\"); fmt.Println(\"two\"); fmt.Println(\"three\"); }",
    );
    assert_eq!(out, vec!["one", "two", "three"]);
}
#[test]
fn string_split_len() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { parts := strings.Split(\"a,b,c\", \",\"); fmt.Println(len(parts)); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn string_join_parts() {
    let out = run_prints(
        "package main; import \"fmt\"; import \"strings\"; func main() { parts := []string{\"a\", \"b\", \"c\"}; fmt.Println(strings.Join(parts, \"-\")); }",
    );
    assert_eq!(out, vec!["a-b-c"]);
}

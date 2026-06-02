use crate::helpers::*;

// ── Control flow: for loops ──────────────────────────────────────────────────

#[test]
fn for_classic_count_up() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { for i := 0; i < 5; i++ { fmt.Println(i); } }",
    );
    assert_eq!(out, vec!["0", "1", "2", "3", "4"]);
}
#[test]
fn for_count_down() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { for i := 3; i > 0; i-- { fmt.Println(i); } }",
    );
    assert_eq!(out, vec!["3", "2", "1"]);
}
#[test]
fn for_step_two() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { for i := 0; i < 10; i = i + 2 { fmt.Println(i); } }",
    );
    assert_eq!(out, vec!["0", "2", "4", "6", "8"]);
}
#[test]
fn for_sum_ten() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := 0; for i := 1; i <= 10; i++ { s = s + i }; fmt.Println(s); }",
    );
    assert_eq!(out, vec!["55"]);
}
#[test]
fn while_style_count() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { n := 1; for n <= 4 { fmt.Println(n); n++ } }",
    );
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}
#[test]
fn infinite_loop_break_at_five() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { i := 0; for { if i == 3 { break }; fmt.Println(i); i++ } }",
    );
    assert_eq!(out, vec!["0", "1", "2"]);
}
#[test]
fn continue_skip_even() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { for i := 0; i < 6; i++ { if i % 2 == 0 { continue }; fmt.Println(i); } }",
    );
    assert_eq!(out, vec!["1", "3", "5"]);
}
#[test]
fn continue_skip_three() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { for i := 0; i < 5; i++ { if i == 3 { continue }; fmt.Println(i); } }",
    );
    assert_eq!(out, vec!["0", "1", "2", "4"]);
}
#[test]
fn nested_for_multiplication() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { for i := 1; i <= 2; i++ { for j := 1; j <= 2; j++ { fmt.Println(i * j); } } }",
    );
    assert_eq!(out, vec!["1", "2", "2", "4"]);
}
#[test]
fn range_over_string_slice() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { langs := []string{\"go\",\"rust\",\"c\"}; for _, l := range langs { fmt.Println(l); } }",
    );
    assert_eq!(out, vec!["go", "rust", "c"]);
}

// ── Control flow: switch ─────────────────────────────────────────────────────

#[test]
fn switch_single_match() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 1; switch x { case 1: fmt.Println(\"one\") } }",
    );
    assert_eq!(out, vec!["one"]);
}
#[test]
fn switch_no_match_no_default() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 5; switch x { case 1: fmt.Println(\"one\") }; fmt.Println(\"done\"); }",
    );
    assert_eq!(out, vec!["done"]);
}
#[test]
fn switch_multiple_cases() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { for i := 1; i <= 3; i++ { switch i { case 1: fmt.Println(\"a\"); case 2: fmt.Println(\"b\"); case 3: fmt.Println(\"c\"); } } }",
    );
    assert_eq!(out, vec!["a", "b", "c"]);
}
#[test]
fn switch_string_expr() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { lang := \"go\"; switch lang { case \"go\": fmt.Println(\"golang\"); case \"rust\": fmt.Println(\"rust-lang\"); default: fmt.Println(\"other\"); } }",
    );
    assert_eq!(out, vec!["golang"]);
}
#[test]
fn switch_expression_in_case() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 4; switch { case x < 0: fmt.Println(\"neg\"); case x == 0: fmt.Println(\"zero\"); case x > 0: fmt.Println(\"pos\"); } }",
    );
    assert_eq!(out, vec!["pos"]);
}
#[test]
fn switch_bool_expr() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { n := 15; switch { case n % 15 == 0: fmt.Println(\"fizzbuzz\"); case n % 3 == 0: fmt.Println(\"fizz\"); case n % 5 == 0: fmt.Println(\"buzz\"); default: fmt.Println(\"n\"); } }",
    );
    assert_eq!(out, vec!["fizzbuzz"]);
}

// ── Control flow: if ─────────────────────────────────────────────────────────

#[test]
fn if_and_condition() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 5; if x > 0 && x < 10 { fmt.Println(\"in range\"); } }",
    );
    assert_eq!(out, vec!["in range"]);
}
#[test]
fn if_or_condition() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 15; if x < 0 || x > 10 { fmt.Println(\"out\"); } }",
    );
    assert_eq!(out, vec!["out"]);
}
#[test]
fn if_not_condition() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { done := false; if !done { fmt.Println(\"not done\"); } }",
    );
    assert_eq!(out, vec!["not done"]);
}
#[test]
fn nested_if_else() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 5; if x > 10 { fmt.Println(\"big\"); } else { if x > 3 { fmt.Println(\"mid\"); } else { fmt.Println(\"small\"); } } }",
    );
    assert_eq!(out, vec!["mid"]);
}
#[test]
fn chained_else_if() {
    let out = run_prints(
        "package main; import \"fmt\"; func classify(n int) string { if n < 0 { return \"neg\" } else if n == 0 { return \"zero\" } else if n < 10 { return \"small\" } else { return \"large\" } } func main() { fmt.Println(classify(-1)); fmt.Println(classify(0)); fmt.Println(classify(5)); fmt.Println(classify(100)); }",
    );
    assert_eq!(out, vec!["neg", "zero", "small", "large"]);
}

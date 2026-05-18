use crate::helpers::*;

#[test] fn variadic_sum_zero() {
    let out = run_prints("package main; import \"fmt\"; func sum(nums ...int) int { total := 0; for _, n := range nums { total = total + n }; return total } func main() { fmt.Println(sum()); }");
    assert_eq!(out, vec!["0"]);
}
#[test] fn variadic_sum_one() {
    let out = run_prints("package main; import \"fmt\"; func sum(nums ...int) int { total := 0; for _, n := range nums { total = total + n }; return total } func main() { fmt.Println(sum(5)); }");
    assert_eq!(out, vec!["5"]);
}
#[test] fn variadic_sum_three() {
    let out = run_prints("package main; import \"fmt\"; func sum(nums ...int) int { total := 0; for _, n := range nums { total = total + n }; return total } func main() { fmt.Println(sum(1, 2, 3)); }");
    assert_eq!(out, vec!["6"]);
}
#[test] fn variadic_sum_five() {
    let out = run_prints("package main; import \"fmt\"; func sum(nums ...int) int { total := 0; for _, n := range nums { total = total + n }; return total } func main() { fmt.Println(sum(1, 2, 3, 4, 5)); }");
    assert_eq!(out, vec!["15"]);
}
#[test] fn variadic_len() {
    let out = run_prints("package main; import \"fmt\"; func count(args ...string) int { return len(args) } func main() { fmt.Println(count(\"a\", \"b\", \"c\")); }");
    assert_eq!(out, vec!["3"]);
}
#[test] fn variadic_with_fixed_param() {
    let out = run_prints("package main; import \"fmt\"; func repeat(prefix string, items ...string) { for _, s := range items { fmt.Println(prefix + s); } } func main() { repeat(\"go:\", \"lang\", \"play\"); }");
    assert_eq!(out, vec!["go:lang", "go:play"]);
}
#[test] fn variadic_max() {
    let out = run_prints("package main; import \"fmt\"; func max(nums ...int) int { m := nums[0]; for _, n := range nums { if n > m { m = n } }; return m } func main() { fmt.Println(max(3, 1, 4, 1, 5, 9, 2)); }");
    assert_eq!(out, vec!["9"]);
}
#[test] fn variadic_concat_strings() {
    let out = run_prints("package main; import \"fmt\"; func joinAll(sep string, parts ...string) string { r := \"\"; i := 0; for _, p := range parts { if i > 0 { r = r + sep }; r = r + p; i++ }; return r } func main() { fmt.Println(joinAll(\"-\", \"a\", \"b\", \"c\")); }");
    assert_eq!(out, vec!["a-b-c"]);
}
#[test] fn multiple_return_add_sub() {
    let out = run_prints("package main; import \"fmt\"; func addSub(a int, b int) (int, int) { return a + b, a - b } func main() { s, d := addSub(10, 3); fmt.Println(s); fmt.Println(d); }");
    assert_eq!(out, vec!["13", "7"]);
}
#[test] fn multiple_return_minmax() {
    let out = run_prints("package main; import \"fmt\"; func minMax(a int, b int) (int, int) { if a < b { return a, b }; return b, a } func main() { lo, hi := minMax(7, 3); fmt.Println(lo); fmt.Println(hi); }");
    assert_eq!(out, vec!["3", "7"]);
}
#[test] fn multiple_return_swap() {
    let out = run_prints("package main; import \"fmt\"; func swap(a int, b int) (int, int) { return b, a } func main() { x, y := swap(1, 2); fmt.Println(x); fmt.Println(y); }");
    assert_eq!(out, vec!["2", "1"]);
}
#[test] fn multiple_return_blank_first() {
    let out = run_prints("package main; import \"fmt\"; func pair() (int, int) { return 5, 10 } func main() { _, b := pair(); fmt.Println(b); }");
    assert_eq!(out, vec!["10"]);
}
#[test] fn multiple_return_three_values() {
    let out = run_prints("package main; import \"fmt\"; func triple() (int, int, int) { return 1, 2, 3 } func main() { a, b, c := triple(); fmt.Println(a + b + c); }");
    assert_eq!(out, vec!["6"]);
}
#[test] fn variadic_bool_all() {
    let out = run_prints("package main; import \"fmt\"; func allTrue(vals ...bool) bool { for _, v := range vals { if !v { return false } }; return true } func main() { fmt.Println(allTrue(true, true, true)); }");
    assert_eq!(out, vec!["true"]);
}
#[test] fn variadic_bool_any_false() {
    let out = run_prints("package main; import \"fmt\"; func allTrue(vals ...bool) bool { for _, v := range vals { if !v { return false } }; return true } func main() { fmt.Println(allTrue(true, false, true)); }");
    assert_eq!(out, vec!["false"]);
}

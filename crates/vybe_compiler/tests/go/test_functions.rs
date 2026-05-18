use crate::helpers::*;

#[test] fn simple_function() {
    let out = run_prints("package main; import \"fmt\"; func greet() { fmt.Println(\"hello\"); } func main() { greet(); }");
    assert_eq!(out, vec!["hello"]);
}
#[test] fn function_with_params() {
    let out = run_prints("package main; import \"fmt\"; func add(a int, b int) { fmt.Println(a + b); } func main() { add(3, 4); }");
    assert_eq!(out, vec!["7"]);
}
#[test] fn function_with_return() {
    let out = run_prints("package main; import \"fmt\"; func add(a int, b int) int { return a + b } func main() { fmt.Println(add(3, 4)); }");
    assert_eq!(out, vec!["7"]);
}
#[test] fn function_multiple_returns() {
    let out = run_prints("package main; import \"fmt\"; func divmod(a int, b int) (int, int) { return a / b, a % b } func main() { q, r := divmod(17, 5); fmt.Println(q); fmt.Println(r); }");
    assert_eq!(out, vec!["3", "2"]);
}
#[test] fn recursive_function() {
    let out = run_prints("package main; import \"fmt\"; func factorial(n int) int { if n <= 1 { return 1 }; return n * factorial(n-1) } func main() { fmt.Println(factorial(5)); }");
    assert_eq!(out, vec!["120"]);
}
#[test] fn closure_literal() {
    let out = run_prints("package main; import \"fmt\"; func main() { f := func(x int) int { return x * 2 }; fmt.Println(f(5)); }");
    assert_eq!(out, vec!["10"]);
}
#[test] fn function_called_twice() {
    let out = run_prints("package main; import \"fmt\"; func say(s string) { fmt.Println(s); } func main() { say(\"hi\"); say(\"bye\"); }");
    assert_eq!(out, vec!["hi", "bye"]);
}
#[test] fn function_bool_param() {
    let out = run_prints("package main; import \"fmt\"; func printBool(b bool) { if b { fmt.Println(\"yes\"); } else { fmt.Println(\"no\"); } } func main() { printBool(true); printBool(false); }");
    assert_eq!(out, vec!["yes", "no"]);
}
#[test] fn function_three_params() {
    let out = run_prints("package main; import \"fmt\"; func sum3(a int, b int, c int) int { return a + b + c } func main() { fmt.Println(sum3(1, 2, 3)); }");
    assert_eq!(out, vec!["6"]);
}
#[test] fn function_string_return() {
    let out = run_prints("package main; import \"fmt\"; func label(n int) string { if n > 0 { return \"pos\" }; return \"nonpos\" } func main() { fmt.Println(label(5)); fmt.Println(label(-1)); }");
    assert_eq!(out, vec!["pos", "nonpos"]);
}
#[test] fn function_early_return() {
    let out = run_prints("package main; import \"fmt\"; func check(n int) string { if n < 0 { return \"neg\" }; return \"ok\" } func main() { fmt.Println(check(-1)); fmt.Println(check(1)); }");
    assert_eq!(out, vec!["neg", "ok"]);
}
#[test] fn higher_order_map_func() {
    let out = run_prints("package main; import \"fmt\"; func mapInts(s []int, f func(int) int) []int { r := []int{}; for _, v := range s { r = append(r, f(v)); }; return r } func main() { doubled := mapInts([]int{1, 2, 3}, func(x int) int { return x * 2 }); for _, v := range doubled { fmt.Println(v); } }");
    assert_eq!(out, vec!["2", "4", "6"]);
}
#[test] fn higher_order_filter() {
    let out = run_prints("package main; import \"fmt\"; func filter(s []int, f func(int) bool) []int { r := []int{}; for _, v := range s { if f(v) { r = append(r, v) } }; return r } func main() { evens := filter([]int{1,2,3,4,5}, func(x int) bool { return x % 2 == 0 }); for _, v := range evens { fmt.Println(v); } }");
    assert_eq!(out, vec!["2", "4"]);
}
#[test] fn higher_order_reduce() {
    let out = run_prints("package main; import \"fmt\"; func reduce(s []int, init int, f func(int, int) int) int { acc := init; for _, v := range s { acc = f(acc, v) }; return acc } func main() { total := reduce([]int{1,2,3,4,5}, 0, func(a int, b int) int { return a + b }); fmt.Println(total); }");
    assert_eq!(out, vec!["15"]);
}
#[test] fn function_no_return() {
    compile_ok("package main; func printMsg(s string) { _ = s } func main() { printMsg(\"hello\"); }");
}
#[test] fn function_calls_another() {
    let out = run_prints("package main; import \"fmt\"; func double(n int) int { return n * 2 } func quadruple(n int) int { return double(double(n)) } func main() { fmt.Println(quadruple(3)); }");
    assert_eq!(out, vec!["12"]);
}
#[test] fn function_string_concat_param() {
    let out = run_prints("package main; import \"fmt\"; func concat(a string, b string) string { return a + b } func main() { fmt.Println(concat(\"foo\", \"bar\")); }");
    assert_eq!(out, vec!["foobar"]);
}
#[test] fn func_returns_func() {
    let out = run_prints("package main; import \"fmt\"; func adder(n int) func(int) int { return func(x int) int { return x + n } } func main() { add10 := adder(10); fmt.Println(add10(5)); }");
    assert_eq!(out, vec!["15"]);
}

use crate::helpers::*;

#[test]
fn closure_basic_call() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { f := func() { fmt.Println(\"hello\"); }; f(); }",
    );
    assert_eq!(out, vec!["hello"]);
}
#[test]
fn closure_with_param() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { double := func(x int) int { return x * 2 }; fmt.Println(double(5)); }",
    );
    assert_eq!(out, vec!["10"]);
}
#[test]
fn closure_with_two_params() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { add := func(a int, b int) int { return a + b }; fmt.Println(add(3, 4)); }",
    );
    assert_eq!(out, vec!["7"]);
}
#[test]
fn closure_returned_from_func() {
    let out = run_prints(
        "package main; import \"fmt\"; func makeAdder(n int) func(int) int { return func(x int) int { return x + n } } func main() { add5 := makeAdder(5); fmt.Println(add5(3)); }",
    );
    assert_eq!(out, vec!["8"]);
}
#[test]
fn closure_captures_var() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { x := 10; f := func() int { return x * 2 }; fmt.Println(f()); }",
    );
    assert_eq!(out, vec!["20"]);
}
#[test]
fn closure_as_argument() {
    let out = run_prints(
        "package main; import \"fmt\"; func apply(f func(int) int, x int) int { return f(x) } func main() { sq := func(n int) int { return n * n }; fmt.Println(apply(sq, 6)); }",
    );
    assert_eq!(out, vec!["36"]);
}
#[test]
fn closure_counter() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { count := 0; inc := func() { count++ }; inc(); inc(); inc(); fmt.Println(count); }",
    );
    assert_eq!(out, vec!["3"]);
}
#[test]
fn closure_in_loop() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { sum := 0; f := func(n int) { sum = sum + n }; i := 1; for i <= 4 { f(i); i++ }; fmt.Println(sum); }",
    );
    assert_eq!(out, vec!["10"]);
}
#[test]
fn immediately_invoked_func() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { result := func(a int, b int) int { return a + b }(10, 20); fmt.Println(result); }",
    );
    assert_eq!(out, vec!["30"]);
}
#[test]
fn closure_multiply() {
    let out = run_prints(
        "package main; import \"fmt\"; func makeMultiplier(factor int) func(int) int { return func(x int) int { return x * factor } } func main() { triple := makeMultiplier(3); fmt.Println(triple(7)); }",
    );
    assert_eq!(out, vec!["21"]);
}
#[test]
fn closure_string_builder() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { prefix := \"Hello\"; greet := func(name string) string { return prefix + \" \" + name }; fmt.Println(greet(\"World\")); }",
    );
    assert_eq!(out, vec!["Hello World"]);
}
#[test]
fn function_as_map_value() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { ops := map[string]func(int, int) int{ \"add\": func(a int, b int) int { return a + b }, \"mul\": func(a int, b int) int { return a * b } }; fmt.Println(ops[\"add\"](3, 4)); }",
    );
    assert_eq!(out, vec!["7"]);
}
#[test]
fn closure_with_bool_return() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { isPositive := func(n int) bool { return n > 0 }; fmt.Println(isPositive(5)); fmt.Println(isPositive(-3)); }",
    );
    assert_eq!(out, vec!["true", "false"]);
}
#[test]
fn closure_composition() {
    let out = run_prints(
        "package main; import \"fmt\"; func compose(f func(int) int, g func(int) int) func(int) int { return func(x int) int { return f(g(x)) } } func main() { addOne := func(x int) int { return x + 1 }; double := func(x int) int { return x * 2 }; doubleAndAdd := compose(addOne, double); fmt.Println(doubleAndAdd(4)); }",
    );
    assert_eq!(out, vec!["9"]);
}
#[test]
fn closure_no_param_no_return() {
    compile_ok("package main; func main() { f := func() {}; f(); }");
}

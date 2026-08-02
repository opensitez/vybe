// vybe-test: go/function_literals_closures/recursive_closure_var_decl_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func main() { var f func(int) int
f = func(n int) int { if n == 0 { return 0 }
return f(n-1) }
_ = f(3) }

// vybe-test: go/function_literals_closures/closure_with_named_return_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func main() { fn := func() (n int) { n = 1
return }
_, _ = fn(), fn() }

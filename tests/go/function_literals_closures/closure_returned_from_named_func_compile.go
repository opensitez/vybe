// vybe-test: go/function_literals_closures/closure_returned_from_named_func_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func mk() func() { return func() {} }
func main() { _ = mk() }

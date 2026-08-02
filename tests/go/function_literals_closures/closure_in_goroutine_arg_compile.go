// vybe-test: go/function_literals_closures/closure_in_goroutine_arg_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func main() { go func() {}() }

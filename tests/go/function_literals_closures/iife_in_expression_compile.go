// vybe-test: go/function_literals_closures/iife_in_expression_compile
// origin: languages/go/tests/go/test_function_literals_closures.rs
// vybe-test-mode: compile

package main
func main() { _ = func(x int) int { return x }(2) }

// vybe-test: go/functions_patterns_extra/variadic_pass_zero_arguments_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func sum(values ...int) int { return len(values) }
func main() { _ = sum() }

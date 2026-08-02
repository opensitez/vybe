// vybe-test: go/functions_patterns_extra/function_type_parameter_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
type transformer func(int) int
func apply(v int, fn transformer) int { return fn(v) }
func main() { _ = apply }

// vybe-test: go/functions_patterns_extra/function_parameter_named_result_same_type_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func transform(v int) (int, int) { return v, v + 1 }
func main() { _, _ = transform(1) }

// vybe-test: go/functions_patterns_extra/function_literal_with_blank_identifier_param_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { fn := func(_ int, v int) int { return v }
_ = fn }

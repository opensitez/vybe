// vybe-test: go/functions_patterns_extra/function_param_shadow_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func use(v int) int { { v := v + 1
_ = v }
return v }
func main() { _ = use(1) }

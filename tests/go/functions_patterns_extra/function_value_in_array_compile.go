// vybe-test: go/functions_patterns_extra/function_value_in_array_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := [1]func(){func() {}}
_ = values }

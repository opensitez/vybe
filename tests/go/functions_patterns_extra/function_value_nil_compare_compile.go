// vybe-test: go/functions_patterns_extra/function_value_nil_compare_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var fn func()
_ = (fn == nil) }

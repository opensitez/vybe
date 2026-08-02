// vybe-test: go/functions_patterns_extra/anonymous_func_in_if_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { if func() bool { return true }() { _ = 1 } }

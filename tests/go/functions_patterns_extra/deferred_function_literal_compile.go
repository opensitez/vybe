// vybe-test: go/functions_patterns_extra/deferred_function_literal_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { defer func() { _ = 1 }() }

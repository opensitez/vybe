// vybe-test: go/functions_patterns_extra/function_array_return_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func build() [2]int { return [2]int{1, 2} }
func main() { _ = build() }

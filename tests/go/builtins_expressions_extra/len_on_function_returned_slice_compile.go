// vybe-test: go/builtins_expressions_extra/len_on_function_returned_slice_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func values() []int { return []int{1, 2, 3} }
func main() { _ = len(values()) }

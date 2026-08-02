// vybe-test: go/builtins_expressions_extra/make_nested_slice_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { values := make([][]int, 2)
_ = values }

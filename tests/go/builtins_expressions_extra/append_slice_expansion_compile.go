// vybe-test: go/builtins_expressions_extra/append_slice_expansion_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { a := []int{1}
b := []int{2, 3}
a = append(a, b...)
_ = a }

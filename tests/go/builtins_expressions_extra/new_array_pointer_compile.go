// vybe-test: go/builtins_expressions_extra/new_array_pointer_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { values := new([3]int)
_ = values }

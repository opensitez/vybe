// vybe-test: go/builtins_expressions_extra/new_named_type_pointer_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
type counter int
func main() { value := new(counter)
_ = value }

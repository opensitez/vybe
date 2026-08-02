// vybe-test: go/method_sets_pointer_value/method_set_pointer_expression_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type T struct{}
func (t *T) M() {}
func main() { _ = (*T).M }

// vybe-test: go/method_values/method_expression_with_pointer_receiver
// origin: languages/go/tests/go/test_method_values.rs
// vybe-test-mode: compile

package main
type T struct{}
func (t *T) M() {}
func main() { f := (*T).M
_ = f }

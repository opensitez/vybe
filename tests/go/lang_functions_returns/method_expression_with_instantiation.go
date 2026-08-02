// vybe-test: go/lang_functions_returns/method_expression_with_instantiation
// origin: languages/go/tests/go/test_lang_functions_returns.rs
// vybe-test-mode: compile

package main
type T int
func (T) M() {}
func main() { _ = T.M }

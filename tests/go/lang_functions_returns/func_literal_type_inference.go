// vybe-test: go/lang_functions_returns/func_literal_type_inference
// origin: languages/go/tests/go/test_lang_functions_returns.rs
// vybe-test-mode: compile

package main
func main() { _ = func(x int) int { return x } }

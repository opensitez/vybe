// vybe-test: go/lang_functions_returns/return_discard_values_compile
// origin: languages/go/tests/go/test_lang_functions_returns.rs
// vybe-test-mode: compile

package main
func f() (int, int) { return 1, 2 }
func main() { f() }

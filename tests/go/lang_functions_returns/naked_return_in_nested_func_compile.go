// vybe-test: go/lang_functions_returns/naked_return_in_nested_func_compile
// origin: languages/go/tests/go/test_lang_functions_returns.rs
// vybe-test-mode: compile

package main
func f() (n int) { n = 1
return }
func main() { _ = f() }

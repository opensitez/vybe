// vybe-test: go/lang_builtins_control/named_result_bare_return
// origin: languages/go/tests/go/test_lang_builtins_control.rs
// vybe-test-mode: compile

package main
func f() (n int) { n = 3
return }
func main() { _ = f() }

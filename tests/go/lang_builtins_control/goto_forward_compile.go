// vybe-test: go/lang_builtins_control/goto_forward_compile
// origin: languages/go/tests/go/test_lang_builtins_control.rs
// vybe-test-mode: compile

package main
func main() { goto L
L: return }

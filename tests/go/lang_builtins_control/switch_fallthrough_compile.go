// vybe-test: go/lang_builtins_control/switch_fallthrough_compile
// origin: languages/go/tests/go/test_lang_builtins_control.rs
// vybe-test-mode: compile

package main
func main() { switch 1 { case 1: fallthrough
case 2: } }

// vybe-test: go/switch_fallthrough_extended/switch_init_undeclared_var_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func main() { switch y; y { case 1: _ = y } }

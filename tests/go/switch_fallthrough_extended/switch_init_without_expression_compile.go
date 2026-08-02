// vybe-test: go/switch_fallthrough_extended/switch_init_without_expression_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func main() { switch x := 1; { case true: _ = x } }

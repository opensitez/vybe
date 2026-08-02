// vybe-test: go/switch_fallthrough_extended/switch_bool_non_constant_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func main() { b := true
switch b { case 1: _ = b } }

// vybe-test: go/switch_fallthrough_extended/switch_duplicate_case_value_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func main() { switch 1 { case 1, 1: _ = 1 } }

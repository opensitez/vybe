// vybe-test: go/switch_fallthrough_extended/switch_case_with_fallthrough_to_same_case_compile
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs
// vybe-test-mode: compile

package main
func main() { switch 1 { case 1: fallthrough
fallthrough
case 2: _ = 1 } }

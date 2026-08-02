// vybe-test: go/control_flow_patterns_extra/switch_with_fallthrough_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { switch 1 { case 1: fallthrough
case 2: } }

// vybe-test: go/control_flow_patterns_extra/label_before_switch_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { Start: switch 1 { case 1: break Start } }

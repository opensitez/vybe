// vybe-test: go/control_flow_patterns_extra/labeled_continue_through_nested_switch_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { Outer: for i := 0; i < 1; i++ { switch i { case 0: continue Outer } } }

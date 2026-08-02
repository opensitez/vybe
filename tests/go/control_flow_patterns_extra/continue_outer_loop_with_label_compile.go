// vybe-test: go/control_flow_patterns_extra/continue_outer_loop_with_label_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { Outer: for i := 0; i < 2; i++ { for j := 0; j < 2; j++ { _ = i + j
continue Outer } } }

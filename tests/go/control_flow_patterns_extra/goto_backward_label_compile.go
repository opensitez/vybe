// vybe-test: go/control_flow_patterns_extra/goto_backward_label_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { i := 0
Loop: i++
if i < 2 { goto Loop } }

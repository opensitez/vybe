// vybe-test: go/control_flow_patterns_extra/range_over_array_pointer_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := &[2]int{1, 2}
for _, v := range values { _ = v } }

// vybe-test: go/control_flow_patterns_extra/range_blank_identifiers_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []int{1, 2}
for _, _ = range values { } }

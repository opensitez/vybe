// vybe-test: go/control_flow_patterns_extra/range_over_map_keys_compile
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1}
for k := range values { _ = k } }

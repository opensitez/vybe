// vybe-test: go/range_iteration_extra/range_over_map_key_value_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1}
for key, value := range values { _, _ = key, value } }

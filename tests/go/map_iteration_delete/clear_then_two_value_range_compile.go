// vybe-test: go/map_iteration_delete/clear_then_two_value_range_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1}
clear(values)
for key, value := range values { _, _ = key, value } }

// vybe-test: go/map_iteration_delete/delete_current_key_during_two_value_range_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1}
for key, value := range values { if value == 1 { delete(values, key) } } }

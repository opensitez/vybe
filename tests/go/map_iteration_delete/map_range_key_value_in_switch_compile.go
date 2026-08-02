// vybe-test: go/map_iteration_delete/map_range_key_value_in_switch_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1}
for key, value := range values { switch key { case "a": _ = value } } }

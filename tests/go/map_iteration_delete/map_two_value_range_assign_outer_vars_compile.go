// vybe-test: go/map_iteration_delete/map_two_value_range_assign_outer_vars_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1}
var key string
var value int
for key, value = range values { _, _ = key, value } }

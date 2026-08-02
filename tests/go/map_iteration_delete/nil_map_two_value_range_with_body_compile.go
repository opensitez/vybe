// vybe-test: go/map_iteration_delete/nil_map_two_value_range_with_body_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { var values map[string]int
for key, value := range values { _, _ = key, value } }

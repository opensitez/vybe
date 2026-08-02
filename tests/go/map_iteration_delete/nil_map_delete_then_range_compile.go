// vybe-test: go/map_iteration_delete/nil_map_delete_then_range_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { var values map[string]int
delete(values, "z")
for key := range values { _ = key } }

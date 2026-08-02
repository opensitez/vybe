// vybe-test: go/map_iteration_delete/delete_during_nested_map_of_slice_range_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { outer := map[string]map[string]int{"x": {"a": 1}}
for _, inner := range outer { for key := range inner { delete(inner, key) } } }

// vybe-test: go/map_iteration_delete/delete_during_range_labeled_break_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1, "b": 2}
outer: for key := range values { delete(values, key)
break outer } }

// vybe-test: go/map_iteration_delete/delete_in_range_with_continue_compile
// origin: languages/go/tests/go/test_map_iteration_delete.rs
// vybe-test-mode: compile

package main
func main() { values := map[int]int{1: 1, 2: 2}
for key, value := range values { if value == 1 { delete(values, key)
continue }
_ = key } }

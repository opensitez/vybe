// vybe-test: go/range_iteration_extra/range_over_map_of_slices_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string][]int{"a": []int{1}}
for _, item := range values { _ = item } }

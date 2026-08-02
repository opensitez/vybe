// vybe-test: go/for_range_extended/range_map_delete_during_iteration_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { m := map[string]int{"a": 1, "b": 2}
for k := range m { delete(m, k)
break } }

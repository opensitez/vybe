// vybe-test: go/for_range_extended/range_map_in_if_condition_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { m := map[int]int{1: 2}
for k := range m { if k > 0 { _ = m[k] } } }

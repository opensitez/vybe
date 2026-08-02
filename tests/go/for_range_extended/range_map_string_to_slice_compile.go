// vybe-test: go/for_range_extended/range_map_string_to_slice_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { m := map[string][]int{"a": {1}}
for k, v := range m { _, _ = k, len(v) } }

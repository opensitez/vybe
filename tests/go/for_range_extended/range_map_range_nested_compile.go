// vybe-test: go/for_range_extended/range_map_range_nested_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { outer := map[string]map[int]bool{"a": {1: true}}
for _, inner := range outer { for k := range inner { _ = k } } }

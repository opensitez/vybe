// vybe-test: go/for_range_extended/range_map_key_only_with_break_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for k := range map[string]int{"a": 1} { _ = k
break } }

// vybe-test: go/range_iteration_extra/range_over_map_literal_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { for key, value := range map[string]int{"a": 1} { _, _ = key, value } }

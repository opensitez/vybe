// vybe-test: go/range_iteration_extra/range_over_nil_map_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { var values map[string]int
for key := range values { _ = key } }

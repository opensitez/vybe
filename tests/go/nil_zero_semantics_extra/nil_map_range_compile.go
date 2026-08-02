// vybe-test: go/nil_zero_semantics_extra/nil_map_range_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var values map[string]int
for k, v := range values { _, _ = k, v } }

// vybe-test: go/for_range_extended/range_nil_map_key_only_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { var m map[string]int
for k := range m { _ = k } }

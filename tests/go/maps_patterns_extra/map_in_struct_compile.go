// vybe-test: go/maps_patterns_extra/map_in_struct_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
type holder struct { values map[string]int }
func main() { _ = holder{} }

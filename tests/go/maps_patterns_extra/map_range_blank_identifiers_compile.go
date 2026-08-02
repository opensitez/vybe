// vybe-test: go/maps_patterns_extra/map_range_blank_identifiers_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{}
for _, value := range values { _ = value } }

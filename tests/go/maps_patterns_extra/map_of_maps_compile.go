// vybe-test: go/maps_patterns_extra/map_of_maps_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = map[string]map[string]int{"a": map[string]int{"b": 1}} }

// vybe-test: go/maps_patterns_extra/map_parameter_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func use(values map[string]int) int { return len(values) }
func main() { _ = use }

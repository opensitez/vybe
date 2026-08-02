// vybe-test: go/maps_patterns_extra/map_return_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func build() map[string]int { return map[string]int{} }
func main() { _ = build() }

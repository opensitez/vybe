// vybe-test: go/maps_patterns_extra/map_of_structs_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int }
func main() { _ = map[string]point{"a": {x: 1}} }

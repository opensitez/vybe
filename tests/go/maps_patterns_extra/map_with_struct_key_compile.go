// vybe-test: go/maps_patterns_extra/map_with_struct_key_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int
y int }
func main() { _ = map[point]int{} }

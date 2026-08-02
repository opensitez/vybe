// vybe-test: go/maps_patterns_extra/map_value_struct_field_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
type point struct { x int }
func main() { values := map[string]point{"a": {x: 1}}
_ = values["a"].x }

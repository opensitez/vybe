// vybe-test: go/maps_keys_values_equal/maps_values_struct_value
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
type S struct { N int }
func main() { _ = maps.Values(map[int]S{1: {N: 1}}) }

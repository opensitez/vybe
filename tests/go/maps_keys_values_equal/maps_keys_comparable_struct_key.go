// vybe-test: go/maps_keys_values_equal/maps_keys_comparable_struct_key
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
type K struct { A int }
func main() { _ = maps.Keys(map[K]string{{1}: "v"}) }

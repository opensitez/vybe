// vybe-test: go/maps_keys_values_equal/maps_keys_after_clear_reassign
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { m := map[int]int{1: 1, 2: 2}
m = map[int]int{}
_ = maps.Keys(m) }

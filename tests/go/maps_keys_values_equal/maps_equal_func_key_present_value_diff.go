// vybe-test: go/maps_keys_values_equal/maps_equal_func_key_present_value_diff
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.EqualFunc(map[string]bool{"a": true}, map[string]bool{"a": false}, func(x, y bool) bool { return x == y }) }

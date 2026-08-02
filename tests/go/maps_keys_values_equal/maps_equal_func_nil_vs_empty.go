// vybe-test: go/maps_keys_values_equal/maps_equal_func_nil_vs_empty
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { var a map[int]int
_ = maps.EqualFunc(a, map[int]int{}, func(x, y int) bool { return x == y }) }

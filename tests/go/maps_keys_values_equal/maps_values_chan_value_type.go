// vybe-test: go/maps_keys_values_equal/maps_values_chan_value_type
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { _ = maps.Values(map[int]chan int{1: make(chan int, 1)}) }

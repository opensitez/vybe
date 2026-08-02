// vybe-test: go/maps_keys_values_equal/maps_values_slice_value_type
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs

package main
import "fmt"
import "maps"
func main() { m := map[int][]int{1: {1, 2}, 2: {3}}
total := 0
for v := range maps.Values(m) { total += len(v) }
fmt.Println(total) }

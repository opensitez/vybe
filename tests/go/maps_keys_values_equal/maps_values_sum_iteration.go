// vybe-test: go/maps_keys_values_equal/maps_values_sum_iteration
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs

package main
import "fmt"
import "maps"
func main() { m := map[int]int{1: 10, 2: 20, 3: 30}
sum := 0
for v := range maps.Values(m) { sum += v }
fmt.Println(sum) }

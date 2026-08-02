// vybe-test: go/maps_keys_values_equal/maps_keys_values_large_map_count
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs

package main
import "fmt"
import "maps"
func main() { m := map[int]int{}
for i := 0; i < 10; i++ { m[i] = i * 10 }
fmt.Println(len(maps.Keys(m)))
fmt.Println(len(maps.Values(m))) }

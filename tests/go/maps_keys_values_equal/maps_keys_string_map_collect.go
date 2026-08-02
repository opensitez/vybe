// vybe-test: go/maps_keys_values_equal/maps_keys_string_map_collect
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs

package main
import "fmt"
import "maps"
func main() { m := map[string]int{"x": 1, "y": 2}
keys := maps.Keys(m)
found := 0
for _, k := range keys { if k == "x" || k == "y" { found++ } }
fmt.Println(found) }

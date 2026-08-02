// vybe-test: go/maps_keys_values_equal/maps_values_string_elements
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs

package main
import "fmt"
import "maps"
func main() { m := map[int]string{1: "go", 2: "vybe"}
count := 0
for v := range maps.Values(m) { if len(v) >= 2 { count++ } }
fmt.Println(count) }

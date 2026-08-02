// vybe-test: go/maps_keys_values_equal/maps_values_bool_values
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs

package main
import "fmt"
import "maps"
func main() { m := map[string]bool{"a": true, "b": false}
trues := 0
for v := range maps.Values(m) { if v { trues++ } }
fmt.Println(trues) }

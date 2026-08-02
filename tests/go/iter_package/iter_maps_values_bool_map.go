// vybe-test: go/iter_package/iter_maps_values_bool_map
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "maps"
func main() { m := map[string]bool{"a": true, "b": false, "c": true}
trues := 0
for v := range maps.Values(m) { if v { trues++ } }
fmt.Println(trues) }

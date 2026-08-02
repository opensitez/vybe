// vybe-test: go/iter_package/iter_maps_values_sum
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "maps"
func main() { m := map[string]int{"a": 1, "b": 2, "c": 3}
sum := 0
for v := range maps.Values(m) { sum += v }
fmt.Println(sum) }

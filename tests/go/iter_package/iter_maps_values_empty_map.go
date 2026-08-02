// vybe-test: go/iter_package/iter_maps_values_empty_map
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "maps"
func main() { count := 0
for range maps.Values(map[int]int{}) { count++ }
fmt.Println(count) }

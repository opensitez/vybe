// vybe-test: go/iter_package/iter_maps_values_nil_map
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "maps"
func main() { var m map[int]int
n := 0
for range maps.Values(m) { n++ }
fmt.Println(n) }

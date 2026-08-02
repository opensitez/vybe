// vybe-test: go/iter_package/iter_maps_values_after_assignment
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "maps"
func main() { m := map[int]int{1: 1}
m[2] = 2
sum := 0
for v := range maps.Values(m) { sum += v }
fmt.Println(sum) }

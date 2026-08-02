// vybe-test: go/iter_package/iter_nested_range_slices_all
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { outer := [][]int{{1, 2}, {3}}
total := 0
for oi := range slices.All(outer) { for ii := range slices.All(outer[oi]) { total += outer[oi][ii] } }
fmt.Println(total) }

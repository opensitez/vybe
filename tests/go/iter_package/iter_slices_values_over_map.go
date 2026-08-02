// vybe-test: go/iter_package/iter_slices_values_over_map
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { m := map[int]int{1: 10, 2: 20}
sum := 0
for v := range slices.Values(m) { sum += v }
fmt.Println(sum) }

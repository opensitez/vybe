// vybe-test: go/iter_package/iter_slices_values_empty
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { n := 0
for range slices.Values(map[string]int{}) { n++ }
fmt.Println(n) }

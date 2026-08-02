// vybe-test: go/iter_package/iter_slices_all_empty_slice
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { s := []int{}
count := 0
for range slices.All(s) { count++ }
fmt.Println(count) }

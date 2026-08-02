// vybe-test: go/iter_package/iter_slices_all_single_element
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { s := []int{42}
idx := -1
for i := range slices.All(s) { idx = i }
fmt.Println(idx) }

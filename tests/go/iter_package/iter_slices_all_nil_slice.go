// vybe-test: go/iter_package/iter_slices_all_nil_slice
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { var s []int
n := 0
for range slices.All(s) { n++ }
fmt.Println(n) }

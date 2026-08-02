// vybe-test: go/iter_package/iter_slices_all_modify_via_index
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { s := []int{1, 2, 3}
for i := range slices.All(s) { s[i] *= 2 }
fmt.Println(s[0])
fmt.Println(s[2]) }

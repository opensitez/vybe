// vybe-test: go/iter_package/iter_slices_all_index_sum
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { s := []int{10, 20, 30}
sum := 0
for i := range slices.All(s) { sum += s[i] }
fmt.Println(sum) }

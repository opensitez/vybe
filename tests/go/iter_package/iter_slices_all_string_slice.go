// vybe-test: go/iter_package/iter_slices_all_string_slice
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { s := []string{"x", "y"}
acc := ""
for i := range slices.All(s) { acc += s[i] }
fmt.Println(acc) }

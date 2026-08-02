// vybe-test: go/iter_package/iter_slices_all_collect_indices
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { s := []string{"a", "b", "c"}
n := 0
for i := range slices.All(s) { if s[i] != "" { n++ } }
fmt.Println(n) }

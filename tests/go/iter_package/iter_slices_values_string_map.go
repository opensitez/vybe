// vybe-test: go/iter_package/iter_slices_values_string_map
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "slices"
func main() { m := map[int]string{1: "go", 2: "lang"}
longest := 0
for v := range slices.Values(m) { if len(v) > longest { longest = len(v) } }
fmt.Println(longest) }

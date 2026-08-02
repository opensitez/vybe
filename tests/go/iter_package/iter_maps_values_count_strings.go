// vybe-test: go/iter_package/iter_maps_values_count_strings
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "maps"
func main() { m := map[int]string{1: "go", 2: "vybe"}
n := 0
for v := range maps.Values(m) { n += len(v) }
fmt.Println(n) }

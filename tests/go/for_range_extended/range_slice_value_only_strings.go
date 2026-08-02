// vybe-test: go/for_range_extended/range_slice_value_only_strings
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for _, word := range []string{"go", "vybe"} { total += len(word) }
fmt.Println(total) }

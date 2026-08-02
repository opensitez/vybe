// vybe-test: go/for_range_extended/range_slice_index_only_triple
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for i := range []int{5, 5, 5} { total += i }
fmt.Println(total) }

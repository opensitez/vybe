// vybe-test: go/for_range_extended/range_slice_index_and_value_pairs
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { sum := 0
for i, v := range []int{10, 20, 30} { sum += i + v }
fmt.Println(sum) }

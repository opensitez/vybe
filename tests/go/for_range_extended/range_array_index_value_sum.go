// vybe-test: go/for_range_extended/range_array_index_value_sum
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for i, v := range [4]int{2, 4, 6, 8} { total += i * v }
fmt.Println(total) }

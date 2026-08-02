// vybe-test: go/range_iteration_extra/range_over_slice_index_sum_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { total := 0
for index := range []int{4, 5, 6} { total += index }
fmt.Println(total)
}

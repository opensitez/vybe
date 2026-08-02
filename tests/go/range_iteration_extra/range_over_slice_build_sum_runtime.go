// vybe-test: go/range_iteration_extra/range_over_slice_build_sum_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { values := make([]int, 3)
values[0], values[1], values[2] = 1, 2, 3
total := 0
for _, value := range values { total += value }
fmt.Println(total)
}

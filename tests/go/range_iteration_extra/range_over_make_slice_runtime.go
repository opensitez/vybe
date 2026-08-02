// vybe-test: go/range_iteration_extra/range_over_make_slice_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { values := make([]int, 2)
values[0], values[1] = 3, 4
total := 0
for _, value := range values { total += value }
fmt.Println(total)
}

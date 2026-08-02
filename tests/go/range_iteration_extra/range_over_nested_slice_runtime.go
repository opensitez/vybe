// vybe-test: go/range_iteration_extra/range_over_nested_slice_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { total := 0
for _, row := range [][]int{{1, 2}, {3}} { total += len(row) }
fmt.Println(total)
}

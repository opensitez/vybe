// vybe-test: go/range_iteration_extra/range_over_slice_value_only_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { total := 0
for _, value := range []int{3, 3, 3} { total += value }
fmt.Println(total)
}

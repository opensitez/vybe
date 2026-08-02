// vybe-test: go/range_iteration_extra/range_over_slice_continue_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { total := 0
for _, value := range []int{1, 2, 3} { if value == 2 { continue }
total += value }
fmt.Println(total)
}

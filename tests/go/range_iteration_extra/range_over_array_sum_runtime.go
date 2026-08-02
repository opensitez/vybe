// vybe-test: go/range_iteration_extra/range_over_array_sum_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { total := 0
for _, value := range [3]int{2, 3, 4} { total += value }
fmt.Println(total)
}

// vybe-test: go/range_iteration_extra/range_with_outer_accumulator_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { total := 1
for _, value := range []int{2, 3} { total *= value }
fmt.Println(total)
}

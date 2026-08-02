// vybe-test: go/range_iteration_extra/range_over_array_values_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { last := 0
for _, value := range [3]int{4, 5, 6} { last = value }
fmt.Println(last)
}

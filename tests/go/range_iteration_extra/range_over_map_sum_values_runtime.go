// vybe-test: go/range_iteration_extra/range_over_map_sum_values_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { values := map[string]int{"a": 2, "b": 5}
total := 0
for _, value := range values { total += value }
fmt.Println(total)
}

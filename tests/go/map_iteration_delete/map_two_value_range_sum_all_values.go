// vybe-test: go/map_iteration_delete/map_two_value_range_sum_all_values
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"a": 10, "b": 20, "c": 30}
total := 0
for _, value := range values { total += value }
fmt.Println(total) }

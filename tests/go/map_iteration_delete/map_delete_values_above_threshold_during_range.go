// vybe-test: go/map_iteration_delete/map_delete_values_above_threshold_during_range
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"low": 1, "mid": 5, "high": 9}
for key, value := range values { if value > 5 { delete(values, key) } }
fmt.Println(len(values))
total := 0
for _, value := range values { total += value }
fmt.Println(total) }

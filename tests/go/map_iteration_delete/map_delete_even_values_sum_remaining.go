// vybe-test: go/map_iteration_delete/map_delete_even_values_sum_remaining
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"a": 1, "b": 2, "c": 3, "d": 4}
for key, value := range values { if value%2 == 0 { delete(values, key) } }
total := 0
for _, value := range values { total += value }
fmt.Println(total) }

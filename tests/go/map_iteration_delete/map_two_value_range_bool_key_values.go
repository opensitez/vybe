// vybe-test: go/map_iteration_delete/map_two_value_range_bool_key_values
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[bool]int{true: 5, false: 7}
total := 0
for _, value := range values { total += value }
fmt.Println(total) }

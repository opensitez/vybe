// vybe-test: go/map_iteration_delete/map_two_value_range_blank_key_value_sum
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[int]int{1: 4, 2: 5, 3: 6}
total := 0
for _, value := range values { total += value }
fmt.Println(total) }

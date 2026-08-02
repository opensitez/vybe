// vybe-test: go/map_iteration_delete/map_two_value_range_count_keys_via_len
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"x": 1, "y": 2, "z": 3}
count := 0
for key, _ := range values { _ = key
count++ }
fmt.Println(count) }

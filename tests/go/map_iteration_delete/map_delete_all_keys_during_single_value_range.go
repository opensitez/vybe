// vybe-test: go/map_iteration_delete/map_delete_all_keys_during_single_value_range
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"a": 1, "b": 2, "c": 3}
for key := range values { delete(values, key) }
fmt.Println(len(values)) }

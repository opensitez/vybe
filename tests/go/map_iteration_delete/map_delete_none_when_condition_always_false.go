// vybe-test: go/map_iteration_delete/map_delete_none_when_condition_always_false
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"x": 1, "y": 2}
for key, value := range values { if value < 0 { delete(values, key) } }
fmt.Println(len(values)) }

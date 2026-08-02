// vybe-test: go/map_iteration_delete/map_delete_matching_value_during_two_value_range
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"a": 1, "b": 2, "c": 2}
for key, value := range values { if value == 2 { delete(values, key) } }
fmt.Println(len(values))
fmt.Println(values["a"]) }

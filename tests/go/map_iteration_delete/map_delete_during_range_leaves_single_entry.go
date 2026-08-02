// vybe-test: go/map_iteration_delete/map_delete_during_range_leaves_single_entry
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"only": 7, "gone": 1}
for key, value := range values { if value == 1 { delete(values, key) } }
fmt.Println(len(values))
fmt.Println(values["only"]) }

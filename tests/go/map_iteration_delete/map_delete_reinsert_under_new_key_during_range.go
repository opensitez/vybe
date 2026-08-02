// vybe-test: go/map_iteration_delete/map_delete_reinsert_under_new_key_during_range
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"old": 0}
for key, value := range values { if value == 0 { delete(values, key)
values["fresh"] = 9 } }
fmt.Println(len(values))
fmt.Println(values["fresh"]) }

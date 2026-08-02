// vybe-test: go/map_iteration_delete/map_delete_named_keys_during_range
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { values := map[string]int{"pre_a": 1, "pre_b": 2, "ok": 3}
for key := range values { if key == "pre_a" || key == "pre_b" { delete(values, key) } }
fmt.Println(len(values))
fmt.Println(values["ok"]) }

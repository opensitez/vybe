// vybe-test: go/slices_maps_advanced/map_keys_iteration_order
// origin: languages/go/tests/go/test_slices_maps_advanced.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { m := map[int]int{1: 1, 2: 2}
for k, _ := range m { _ = k } }

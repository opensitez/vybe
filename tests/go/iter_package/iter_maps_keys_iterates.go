// vybe-test: go/iter_package/iter_maps_keys_iterates
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { m := map[int]int{1: 1, 2: 2}
for k := range maps.Keys(m) { _ = k } }

// vybe-test: go/iter_package/iter_maps_all_key_value
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { m := map[int]string{1: "a"}
for k, v := range maps.All(m) { _, _ = k, v } }

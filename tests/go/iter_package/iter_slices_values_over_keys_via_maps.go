// vybe-test: go/iter_package/iter_slices_values_over_keys_via_maps
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "maps"
import "slices"
func main() { m := map[string]int{"a": 1}
for range slices.Values(m) {} }

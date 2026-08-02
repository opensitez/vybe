// vybe-test: go/iter_package/iter_slices_values_struct_map
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "slices"
type P struct { N int }
func main() { m := map[string]P{"a": {1}}
for range slices.Values(m) {} }

// vybe-test: go/iter_package/iter_slices_all_over_bytes
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { b := []byte{'a', 'b'}
for i := range slices.All(b) { _ = b[i] } }

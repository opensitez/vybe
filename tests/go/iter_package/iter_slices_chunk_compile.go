// vybe-test: go/iter_package/iter_slices_chunk_compile
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "slices"
func main() { s := []int{1, 2, 3, 4}
for c := range slices.Chunk(s, 2) { _ = c } }

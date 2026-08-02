// vybe-test: go/slices_maps_stdlib/maps_copy_disjoint_key_sets
// origin: languages/go/tests/go/test_slices_maps_stdlib.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { dst := map[int]bool{}
src := map[int]bool{1: true, 2: false}
_ = maps.Copy(dst, src) }

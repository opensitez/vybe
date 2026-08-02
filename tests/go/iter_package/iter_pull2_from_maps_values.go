// vybe-test: go/iter_package/iter_pull2_from_maps_values
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
import "maps"
func main() { m := map[int]string{1: "a"}
next, stop := iter.Pull2(maps.All(m))
defer stop()
_, _, _ = next() }

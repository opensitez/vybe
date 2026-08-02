// vybe-test: go/iter_package/iter_pull_from_slices_all
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
import "slices"
func main() { s := []int{1, 2}
next, stop := iter.Pull(slices.All(s))
defer stop()
_, _ = next() }

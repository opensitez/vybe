// vybe-test: go/iter_package/iter_pull_from_custom_seq
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
func nums() iter.Seq[int] { return func(yield func(int) bool) { yield(1)
yield(2) } }
func main() { next, stop := iter.Pull(nums())
defer stop()
_ = next }

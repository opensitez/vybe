// vybe-test: go/iter_package/iter_pull_from_empty_seq
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
func main() { seq := func(yield func(int) bool) {}
next, stop := iter.Pull(seq)
defer stop()
_, _ = next() }

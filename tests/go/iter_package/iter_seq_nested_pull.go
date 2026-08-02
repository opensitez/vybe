// vybe-test: go/iter_package/iter_seq_nested_pull
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
func main() { outer := func(yield func(iter.Seq[int]) bool) { yield(func(yield func(int) bool) { yield(1) }) }
for inner := range outer { next, stop := iter.Pull(inner)
stop() } }

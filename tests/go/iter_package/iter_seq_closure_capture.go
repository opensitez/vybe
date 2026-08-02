// vybe-test: go/iter_package/iter_seq_closure_capture
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
func main() { base := 10
seq := func(yield func(int) bool) { yield(base) }
for v := range seq { _ = v } }

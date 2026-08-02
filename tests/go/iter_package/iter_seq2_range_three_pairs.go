// vybe-test: go/iter_package/iter_seq2_range_three_pairs
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
func main() { seq := func(yield func(int, int) bool) { yield(1, 10)
yield(2, 20)
yield(3, 30) }
for k, v := range seq { _, _ = k, v } }

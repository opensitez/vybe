// vybe-test: go/iter_package/iter_seq_type_as_func_value
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
func main() { var seq iter.Seq[int] = func(yield func(int) bool) { yield(1) }
_ = seq }

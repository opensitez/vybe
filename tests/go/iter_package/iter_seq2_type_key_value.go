// vybe-test: go/iter_package/iter_seq2_type_key_value
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "iter"
func main() { var seq iter.Seq2[string, int] = func(yield func(string, int) bool) { yield("k", 1) }
_ = seq }

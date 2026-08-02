// vybe-test: go/iter_package/iter_seq_yield_returns_bool
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
func main() { seq := func(yield func(int) bool) bool { return yield(1) }
_ = seq }

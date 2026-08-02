// vybe-test: go/nil_zero_semantics_extra/zero_value_make_after_nil_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var values []int
values = make([]int, 2)
_ = values }

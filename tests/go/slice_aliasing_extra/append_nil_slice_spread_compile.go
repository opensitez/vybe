// vybe-test: go/slice_aliasing_extra/append_nil_slice_spread_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { var values []int
extra := []int{1, 2}
values = append(values, extra...)
_ = values }

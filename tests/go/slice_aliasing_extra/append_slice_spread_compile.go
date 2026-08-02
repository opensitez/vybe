// vybe-test: go/slice_aliasing_extra/append_slice_spread_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []int{1}
extra := []int{2, 3}
values = append(values, extra...)
_ = values }

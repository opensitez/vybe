// vybe-test: go/slice_aliasing_extra/full_slice_expression_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []int{1, 2, 3}
_ = values[0:2:3] }

// vybe-test: go/slice_aliasing_extra/make_nested_slice_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { values := make([][]int, 2)
_ = values }

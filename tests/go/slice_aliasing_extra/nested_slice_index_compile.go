// vybe-test: go/slice_aliasing_extra/nested_slice_index_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { values := [][]int{{1}}
_ = values[0][0] }

// vybe-test: go/slice_aliasing_extra/nested_slice_literal_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = [][]int{{1, 2}, {3, 4}} }

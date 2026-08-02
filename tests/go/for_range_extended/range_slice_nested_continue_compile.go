// vybe-test: go/for_range_extended/range_slice_nested_continue_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for _, row := range [][]int{{1, 2}, {3}} { for _, v := range row { if v == 2 { continue }
_ = v } } }

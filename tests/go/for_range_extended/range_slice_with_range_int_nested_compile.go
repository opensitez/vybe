// vybe-test: go/for_range_extended/range_slice_with_range_int_nested_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for i := range 2 { for _, v := range []int{i, i + 1} { _ = v } } }

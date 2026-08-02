// vybe-test: go/blank_identifier_extended/blank_range_slice_index_discard_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { for _, v := range []int{1, 2} { _ = v } }

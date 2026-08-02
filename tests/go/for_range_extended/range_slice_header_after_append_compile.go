// vybe-test: go/for_range_extended/range_slice_header_after_append_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { s := []int{1}
s = append(s, 2)
for i, v := range s { _, _ = i, v } }

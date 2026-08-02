// vybe-test: go/for_range_extended/range_over_named_slice_type_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
type digits []int
func main() { d := digits{1, 2}
for i, v := range d { _, _ = i, v } }

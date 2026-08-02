// vybe-test: go/for_range_extended/range_slice_assign_existing_index_value_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { s := []int{1, 2}
var i int
var v int
for i, v = range s { _, _ = i, v } }

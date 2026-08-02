// vybe-test: go/for_range_extended/range_over_interface_slice_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { items := []interface{}{1, "x"}
for i, v := range items { _, _ = i, v } }

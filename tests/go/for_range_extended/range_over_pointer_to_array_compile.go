// vybe-test: go/for_range_extended/range_over_pointer_to_array_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { arr := &[3]int{1, 2, 3}
for i, v := range arr { _, _ = i, v } }

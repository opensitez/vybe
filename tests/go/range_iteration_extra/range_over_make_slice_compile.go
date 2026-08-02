// vybe-test: go/range_iteration_extra/range_over_make_slice_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { values := make([]int, 2)
for _, value := range values { _ = value } }

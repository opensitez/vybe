// vybe-test: go/range_iteration_extra/range_index_only_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []int{1}
for index := range values { _ = index } }

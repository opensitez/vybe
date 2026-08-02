// vybe-test: go/range_iteration_extra/range_assign_existing_vars_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []int{1}
var index int
var value int
for index, value = range values { _, _ = index, value } }

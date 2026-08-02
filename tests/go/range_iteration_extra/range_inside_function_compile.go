// vybe-test: go/range_iteration_extra/range_inside_function_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func use(values []int) int { total := 0
for _, value := range values { total += value }
return total }
func main() { _ = use }

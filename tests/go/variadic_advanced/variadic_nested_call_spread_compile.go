// vybe-test: go/variadic_advanced/variadic_nested_call_spread_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func inner(nums ...int) int { return len(nums) }
func outer(nums ...int) int { return inner(append([]int{0}, nums...)...) }
func main() { _ = outer(2, 3) }

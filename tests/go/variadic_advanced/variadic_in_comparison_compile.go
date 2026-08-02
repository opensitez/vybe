// vybe-test: go/variadic_advanced/variadic_in_comparison_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func lenInts(nums ...int) int { return len(nums) }
func main() { _ = lenInts(1, 2) == 2 }

// vybe-test: go/variadic_advanced/variadic_mixed_types_separate_funcs_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func ints(nums ...int) int { return len(nums) }
func strs(parts ...string) int { return len(parts) }
func main() { _ = ints(1)
_ = strs("x") }

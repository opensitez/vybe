// vybe-test: go/variadic_advanced/variadic_assign_to_slice_var_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func collect(nums ...int) []int { return nums }
func main() { s := collect(1, 2)
_ = s[0] }

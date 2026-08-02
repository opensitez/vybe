// vybe-test: go/variadic_advanced/variadic_named_type_slice_spread_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
type digits []int
func sum(nums ...int) int { return len(nums) }
func main() { d := digits{1, 2}
_ = sum([]int(d)...)}

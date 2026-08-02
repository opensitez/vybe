// vybe-test: go/variadic_advanced/spread_requires_slice_type_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func sink(nums ...int) int { return len(nums) }
func main() { _ = sink([]int{1, 2}...) }

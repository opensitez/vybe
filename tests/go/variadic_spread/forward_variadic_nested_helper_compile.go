// vybe-test: go/variadic_spread/forward_variadic_nested_helper_compile
// origin: languages/go/tests/go/test_variadic_spread.rs
// vybe-test-mode: compile

package main
func sink(nums ...int) int { return len(nums) }
func relay(nums ...int) int { return func(v ...int) int { return sink(v...) }(nums...) }
func main() { _ = relay(1, 2) }

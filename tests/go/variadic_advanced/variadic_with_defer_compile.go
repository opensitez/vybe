// vybe-test: go/variadic_advanced/variadic_with_defer_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func log(nums ...int) int { defer func() { _ = len(nums) }()
return len(nums) }
func main() { _ = log(1, 2) }

// vybe-test: go/variadic_advanced/forward_variadic_identity_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func id(nums ...int) []int { return nums }
func wrap(nums ...int) []int { return id(nums...) }
func main() { _ = wrap(1) }

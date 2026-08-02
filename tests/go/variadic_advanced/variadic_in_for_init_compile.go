// vybe-test: go/variadic_advanced/variadic_in_for_init_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func size(nums ...int) int { return len(nums) }
func main() { for i := size(1, 2, 3); i > 0; i-- { break } }

// vybe-test: go/variadic_advanced/variadic_returned_from_factory_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func build() func(...int) int { return func(nums ...int) int { return len(nums) } }
func main() { f := build()
_ = f(1) }

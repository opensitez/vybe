// vybe-test: go/variadic_spread/forward_variadic_closure_compile
// origin: languages/go/tests/go/test_variadic_spread.rs
// vybe-test-mode: compile

package main
func build() func(...int) int { return func(nums ...int) int { return len(nums) } }
func main() { fn := build()
_ = fn(1, 2, 3) }

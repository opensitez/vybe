// vybe-test: go/variadic_advanced/variadic_spread_from_function_return_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func batch() []int { return []int{1, 2} }
func sum(nums ...int) int { t := 0
for _, n := range nums { t += n }
return t }
func main() { _ = sum(batch()...) }

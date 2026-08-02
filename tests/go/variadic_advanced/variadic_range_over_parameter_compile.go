// vybe-test: go/variadic_advanced/variadic_range_over_parameter_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func sum(nums ...int) int { t := 0
for _, n := range nums { t += n }
return t }
func main() { _ = sum(1, 2, 3) }

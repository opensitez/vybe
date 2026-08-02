// vybe-test: go/variadic_advanced/variadic_in_return_statement_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func max(nums ...int) int { m := nums[0]
for _, n := range nums { if n > m { m = n } }
return m }
func main() { _ = max(3, 9, 1) }

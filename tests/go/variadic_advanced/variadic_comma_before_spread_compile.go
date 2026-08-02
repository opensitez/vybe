// vybe-test: go/variadic_advanced/variadic_comma_before_spread_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func pick(prefix int, rest ...int) int { return prefix + len(rest) }
func main() { extra := []int{4, 5}
_ = pick(1, extra...) }

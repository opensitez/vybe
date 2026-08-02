// vybe-test: go/variadic_advanced/variadic_multiple_spread_sources_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func merge(a ...int) int { return len(a) }
func main() { first := []int{1}
second := []int{2, 3}
_ = merge(append(first, second...)...) }

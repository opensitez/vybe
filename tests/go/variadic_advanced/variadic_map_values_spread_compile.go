// vybe-test: go/variadic_advanced/variadic_map_values_spread_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func keys(m map[string]int) int { return len(m) }
func main() { _ = keys(map[string]int{"a": 1}) }

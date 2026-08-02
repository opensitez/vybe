// vybe-test: go/functions_patterns_extra/higher_order_returning_variadic_compile
// origin: languages/go/tests/go/test_functions_patterns_extra.rs
// vybe-test-mode: compile

package main
func build() func(...int) int { return func(values ...int) int { return len(values) } }
func main() { _ = build() }

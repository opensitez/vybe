// vybe-test: go/function_types_advanced/variadic_func_type_distinct_from_fixed_compile
// origin: languages/go/tests/go/test_function_types_advanced.rs
// vybe-test-mode: compile

package main
type Fixed func(int) int
type Variadic func(...int) int
func useFixed(f Fixed) int { return f(1) }
func useVariadic(v Variadic) int { return v(1, 2) }
func main() { _ = useFixed(func(v int) int { return v })
_ = useVariadic(func(values ...int) int { return len(values) }) }

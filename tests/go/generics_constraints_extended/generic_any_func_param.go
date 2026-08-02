// vybe-test: go/generics_constraints_extended/generic_any_func_param
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func Apply[T any](f func(T) T, v T) T { return f(v) }
func main() { _ = Apply(func(x int) int { return x + 1 }, 1) }

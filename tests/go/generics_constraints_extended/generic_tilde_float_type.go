// vybe-test: go/generics_constraints_extended/generic_tilde_float_type
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
type Real float64
func Square[T ~float64](v T) T { return v * v }
func main() { _ = Square(Real(2)) }

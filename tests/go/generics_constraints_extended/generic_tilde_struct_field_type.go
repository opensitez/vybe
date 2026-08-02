// vybe-test: go/generics_constraints_extended/generic_tilde_struct_field_type
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
type Score int
func Double[T ~int](v T) T { return v * 2 }
func main() { _ = Double(Score(3)) }

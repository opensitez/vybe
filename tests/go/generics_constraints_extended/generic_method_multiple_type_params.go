// vybe-test: go/generics_constraints_extended/generic_method_multiple_type_params
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
type Pair[A, B any] struct { A A
B B }
func (p Pair[A, B]) Swap() Pair[B, A] { return Pair[B, A]{A: p.B, B: p.A} }
func main() { _ = Pair[int, string]{1, "x"}.Swap() }

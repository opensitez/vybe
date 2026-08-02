// vybe-test: go/generics_constraints_extended/generic_ordered_uint_constraint
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
import "cmp"
func MinU[T cmp.Ordered](a, b T) T { if a < b { return a }
return b }
func main() { _ = MinU(uint(1), uint(2)) }

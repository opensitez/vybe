// vybe-test: go/generics_functions/generic_ordered_constraint
// origin: languages/go/tests/go/test_generics_functions.rs
// vybe-test-mode: compile

package main
import "cmp"
func Clamp[T cmp.Ordered](v, lo, hi T) T { if v < lo { return lo }
if v > hi { return hi }
return v }
func main() { _ = Clamp(3, 1, 5) }

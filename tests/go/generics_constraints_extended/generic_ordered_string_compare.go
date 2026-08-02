// vybe-test: go/generics_constraints_extended/generic_ordered_string_compare
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
import "cmp"
func Sorted[T cmp.Ordered](a, b T) bool { return cmp.Less(a, b) }
func main() { _ = Sorted("a", "b") }

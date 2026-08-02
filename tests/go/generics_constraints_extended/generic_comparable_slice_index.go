// vybe-test: go/generics_constraints_extended/generic_comparable_slice_index
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func Index[T comparable](s []T, v T) int { for i, x := range s { if x == v { return i } }
return -1 }
func main() { _ = Index([]string{"a"}, "a") }

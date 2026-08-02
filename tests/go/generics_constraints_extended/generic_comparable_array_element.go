// vybe-test: go/generics_constraints_extended/generic_comparable_array_element
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func HasArray[T comparable](s []T, v T) bool { for _, x := range s { if x == v { return true } }
return false }
func main() { _ = HasArray([][1]int{{1}}, [1]int{1}) }

// vybe-test: go/generics_constraints_extended/generic_method_comparable_contains
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Set[T comparable] struct { items []T }
func (s *Set[T]) Add(v T) { s.items = append(s.items, v) }
func (s Set[T]) Has(v T) bool { for _, x := range s.items { if x == v { return true } }
return false }
func main() { var st Set[int]
st.Add(7)
fmt.Println(st.Has(7))
fmt.Println(st.Has(8)) }

// vybe-test: go/generics_constraints_extended/generic_method_on_constraint_interface
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
import "cmp"
type Sorter[T cmp.Ordered] interface { Sort() }
type Ints []int
func (s Ints) Sort() { for i := 0; i < len(s); i++ { for j := i+1; j < len(s); j++ { if s[j] < s[i] { s[i], s[j] = s[j], s[i] } } } }
func Run[T cmp.Ordered, S Sorter[T]](s S) { s.Sort() }
func main() { data := Ints{3, 1, 2}
Run(data) }

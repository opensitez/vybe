// vybe-test: go/generics_constraints_extended/generic_method_ordered_max_in_slice
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
type Stats[T cmp.Ordered] struct { data []T }
func (s Stats[T]) Max() T { m := s.data[0]
for _, v := range s.data[1:] { if cmp.Less(m, v) { m = v } }
return m }
func main() { fmt.Println(Stats[int]{data: []int{1, 9, 3}}.Max()) }

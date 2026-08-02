// vybe-test: go/generics_types/generic_set_comparable_membership
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Set[T comparable] struct { m map[T]struct{} }
func (s *Set[T]) Add(v T) { if s.m == nil { s.m = make(map[T]struct{}) }
s.m[v] = struct{}{} }
func (s *Set[T]) Has(v T) bool { _, ok := s.m[v]
return ok }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s Set[int]
s.Add(3)
__check(fmt.Sprint(s.Has(3)), "true")
__check(fmt.Sprint(s.Has(4)), "false") }

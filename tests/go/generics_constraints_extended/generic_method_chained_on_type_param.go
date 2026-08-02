// vybe-test: go/generics_constraints_extended/generic_method_chained_on_type_param
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Num[T ~int] struct { V T }
func (n Num[T]) Inc() Num[T] { n.V++
return n }
func (n Num[T]) Value() T { return n.V }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Num[int]{V: 1}.Inc().Value()), "2") }

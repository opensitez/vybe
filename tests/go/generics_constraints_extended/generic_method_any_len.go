// vybe-test: go/generics_constraints_extended/generic_method_any_len
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Bag[T any] struct { items []T }
func (b Bag[T]) Len() int { return len(b.items) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Bag[string]{items: []string{"a", "b"}}.Len()), "2") }

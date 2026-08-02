// vybe-test: go/generics_types/generic_embed_ordered_in_constraint
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
import "cmp"
type Ordered = cmp.Ordered
type Sorter[T Ordered] struct{}
func (Sorter[T]) IsLess(a, b T) bool { return a < b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Sorter[int]{}.IsLess(1, 3)), "true") }

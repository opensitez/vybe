// vybe-test: go/generics_types/generic_union_constraint_float
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Real interface { ~int | ~float64 }
type Half[T Real] struct{}
func (Half[T]) Of(v T) T { return v / 2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Half[float64]{}.Of(5.0)), "2.5") }

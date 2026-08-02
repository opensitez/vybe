// vybe-test: go/generics_types/generic_signed_tilde_constraint
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Signed interface { ~int | ~int64 }
type Abs[T Signed] struct{}
func (Abs[T]) Negate(v T) T { return -v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Abs[int]{}.Negate(7)), "-7") }

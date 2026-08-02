// vybe-test: go/generics_types/generic_calculator_twice_method
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Numeric interface { ~int | ~float64 }
type Calculator[T Numeric] struct{}
func (Calculator[T]) Twice(v T) T { return v + v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Calculator[int]{}.Twice(6)), "12") }

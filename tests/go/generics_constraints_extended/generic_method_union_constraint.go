// vybe-test: go/generics_constraints_extended/generic_method_union_constraint
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Converter[T int | float64] struct { Factor T }
func (c Converter[T]) Scale(v T) T { return v * c.Factor }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Converter[int]{Factor: 3}.Scale(4)), "12") }

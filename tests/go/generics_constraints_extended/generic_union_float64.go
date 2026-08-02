// vybe-test: go/generics_constraints_extended/generic_union_float64
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Double[T float32 | float64](v T) T { return v * 2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Double(2.5)), "5") }

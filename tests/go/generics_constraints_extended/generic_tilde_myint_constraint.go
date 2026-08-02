// vybe-test: go/generics_constraints_extended/generic_tilde_myint_constraint
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type MyInt int
func AddOne[T ~int](v T) T { return v + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(AddOne(MyInt(4))), "5") }

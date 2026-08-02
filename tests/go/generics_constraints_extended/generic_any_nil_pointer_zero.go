// vybe-test: go/generics_constraints_extended/generic_any_nil_pointer_zero
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func ZeroPtr[T any]() *T { return nil }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ZeroPtr[int]() == nil), "true") }

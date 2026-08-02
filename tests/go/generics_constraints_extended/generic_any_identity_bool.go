// vybe-test: go/generics_constraints_extended/generic_any_identity_bool
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func ID[T any](v T) T { return v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ID(true)), "true") }

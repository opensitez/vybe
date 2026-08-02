// vybe-test: go/generics_constraints_extended/generic_tilde_mystring_upper
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "strings"
type Label string
func Upper[T ~string](s T) T { return T(strings.ToUpper(string(s))) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Upper(Label("go"))), "GO") }

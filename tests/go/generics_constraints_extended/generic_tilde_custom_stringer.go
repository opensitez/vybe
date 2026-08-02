// vybe-test: go/generics_constraints_extended/generic_tilde_custom_stringer
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type MyString string
func Quote[T ~string](s T) string { return "\"" + string(s) + "\"" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Quote(MyString("go"))), "\"go\"") }

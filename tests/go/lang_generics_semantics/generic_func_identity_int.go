// vybe-test: go/lang_generics_semantics/generic_func_identity_int
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
func ID[T any](v T) T { return v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ID(7)), "7") }

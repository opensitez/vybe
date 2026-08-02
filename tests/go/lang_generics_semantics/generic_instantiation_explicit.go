// vybe-test: go/lang_generics_semantics/generic_instantiation_explicit
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
func Zero[T any]() T { var z T
return z }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Zero[int]()), "0") }

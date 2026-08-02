// vybe-test: go/lang_generics_semantics/generic_func_two_type_params
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
func Pair[A any, B any](a A, b B) (A, B) { return a, b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x, y := Pair(1, "x")
__check(fmt.Sprint(x), "1")
__check(fmt.Sprint(y), "x") }

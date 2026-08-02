// vybe-test: go/lang_declarations_types/function_type_nil_eq
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var f func(int) int
__check(fmt.Sprint(f == nil), "true") }

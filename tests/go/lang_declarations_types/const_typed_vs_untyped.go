// vybe-test: go/lang_declarations_types/const_typed_vs_untyped
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
const x int = 3
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(x), "3") }

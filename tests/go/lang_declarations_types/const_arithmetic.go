// vybe-test: go/lang_declarations_types/const_arithmetic
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
const x = 10 * 5
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(x), "50") }

// vybe-test: go/lang_expressions/new_returns_pointer
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := new(int)
*p = 6
__check(fmt.Sprint(*p), "6") }

// vybe-test: go/lang_expressions/unary_plus_minus
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x := 5
__check(fmt.Sprint(-x), "-5")
__check(fmt.Sprint(+x), "5") }

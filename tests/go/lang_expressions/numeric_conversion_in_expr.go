// vybe-test: go/lang_expressions/numeric_conversion_in_expr
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var b byte = 65
__check(fmt.Sprint(string(b)), "A") }

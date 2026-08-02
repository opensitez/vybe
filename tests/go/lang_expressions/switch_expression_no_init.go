// vybe-test: go/lang_expressions/switch_expression_no_init
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch 2 { case 2: __check(fmt.Sprint("y"), "y") } }

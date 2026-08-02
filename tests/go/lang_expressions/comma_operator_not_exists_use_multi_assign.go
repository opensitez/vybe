// vybe-test: go/lang_expressions/comma_operator_not_exists_use_multi_assign
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a, b := 1, 2
__check(fmt.Sprint(a+b), "3") }

// vybe-test: go/lang_expressions/type_assertion_single_value
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var i interface{} = 1
__check(fmt.Sprint(i.(int)), "1") }

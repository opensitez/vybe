// vybe-test: go/lang_expressions/dereference_pointer
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x := 2
p := &x
*p = 7
__check(fmt.Sprint(x), "7") }

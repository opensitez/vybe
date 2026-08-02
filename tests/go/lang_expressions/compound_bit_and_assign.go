// vybe-test: go/lang_expressions/compound_bit_and_assign
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := 7
n &= 3
__check(fmt.Sprint(n), "3") }

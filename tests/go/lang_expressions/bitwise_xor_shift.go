// vybe-test: go/lang_expressions/bitwise_xor_shift
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(10 ^ 12), "6")
__check(fmt.Sprint(1 << 3), "8") }

// vybe-test: go/lang_expressions/make_slice_len
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := make([]int, 3)
__check(fmt.Sprint(len(s)), "3") }

// vybe-test: go/lang_expressions/append_returns_same_slice_header
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1}
t := append(s, 2)
__check(fmt.Sprint(t[1]), "2") }

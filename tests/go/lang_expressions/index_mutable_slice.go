// vybe-test: go/lang_expressions/index_mutable_slice
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1,2,3}
s[1] = 9
__check(fmt.Sprint(s[1]), "9") }

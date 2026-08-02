// vybe-test: go/lang_declarations_types/zero_value_int
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var n int
__check(fmt.Sprint(n), "0") }

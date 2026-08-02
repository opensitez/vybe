// vybe-test: go/lang_builtins_control/iota_in_const_block
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
const ( A = iota; B; C )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(C), "2") }

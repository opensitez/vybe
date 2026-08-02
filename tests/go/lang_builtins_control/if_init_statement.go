// vybe-test: go/lang_builtins_control/if_init_statement
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { if x := 3; x > 1 { __check(fmt.Sprint(x), "3") } }

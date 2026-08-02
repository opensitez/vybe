// vybe-test: go/lang_builtins_control/short_decl_shadow_outer
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := 1
{ a := 2
__check(fmt.Sprint(a), "2") } }

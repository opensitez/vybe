// vybe-test: go/lang_builtins_control/panic_recover_value
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer func() { if r := recover(); r != nil { __check(fmt.Sprint(r), "boom") } }()
panic("boom") }

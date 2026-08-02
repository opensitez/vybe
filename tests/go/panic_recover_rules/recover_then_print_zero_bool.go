// vybe-test: go/panic_recover_rules/recover_then_print_zero_bool
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { recover()
var b bool
__check(fmt.Sprint(b), "false") }()
panic("z") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }

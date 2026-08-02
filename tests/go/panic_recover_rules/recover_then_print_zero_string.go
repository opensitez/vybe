// vybe-test: go/panic_recover_rules/recover_then_print_zero_string
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { recover()
var s string
__check(fmt.Sprint(s == ""), "true") }()
panic(1) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }

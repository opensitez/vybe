// vybe-test: go/panic_recover_rules/recover_then_print_zero_float
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { recover()
var f float64
__check(fmt.Sprint(f), "0") }()
panic(true) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
